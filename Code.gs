// Google Docs Hashtag Indexing Script with State Management
//
// This script indexes hashtags in a Google Doc and creates a "Tags" section
// with references to all tagged content.
//
// STATE MANAGEMENT:
// To handle large documents that may timeout (6 minute limit), this script
// implements a resumable execution model:
//
// 1. Two phases: GATHERING (collect hashtags) and WRITING (build Tags section)
// 2. Runtime is tracked; after ~4.5 minutes, state is saved and execution stops
// 3. State is stored in:
//    - JSON file in Drive: tracks phase and progress indices
//    - JSON data file in Drive: stores collected tagChildren as lightweight JSON
//      (much faster than Docs API - elements serialized to text/metadata)
// 4. On next run, if state file is newer than document, execution resumes
// 5. On successful completion, state files are cleaned up
//
// GATHERING PHASE:
// - Iterates through document children, tracking childIndex
// - Collects hashtags and associated content into tagChildren structure
// - If time limit reached, saves state (serializing tagChildren to JSON) and exits
// - On resume, continues from saved childIndex
//
// WRITING PHASE:
// - Iterates through collected tags, tracking currentTagIndex and currentTagChildIndex
// - Writes tag sections to the document
// - If time limit reached, saves state and exits (tagChildren already in JSON file)
// - On resume, continues from saved indices
//
// PERFORMANCE NOTES:
// - MAX_RUNTIME_MS set to 4.5 minutes to allow time for state serialization
// - JSON blob serialization is much faster than Docs API writes
// - Document elements converted to lightweight format (text + metadata only)
//

const TAGS_REGEX_STRING = "#[^\\s]+"
const TAGS_REGEX = new RegExp(TAGS_REGEX_STRING, 'g')
const TAGS_SECTION_NAME = "Tags"
const MAX_CHARS_PER_ENTRY = 150
const ELLIPSIS = "..."
const SAVE_THRESHOLD = 100
const MAX_RUNTIME_MS = 4.5 * 60 * 1000 // 4.5 minutes in milliseconds (save early to allow time for state serialization)
const STATE_FILE_PREFIX = "hashtag_indexing_state"
const TEMP_DATA_PREFIX = "hashtag_temp_data"

function getStateFile_(docId) {
  const fileName = `${STATE_FILE_PREFIX}_${docId}.json`
  const files = DriveApp.getFilesByName(fileName)
  if (files.hasNext()) {
    return files.next()
  }
  return null
}

function getTempDataFile_(docId) {
  const fileName = `${TEMP_DATA_PREFIX}_${docId}.json`
  const files = DriveApp.getFilesByName(fileName)
  if (files.hasNext()) {
    return files.next()
  }
  return null
}

function createStateFile_(docId) {
  const fileName = `${STATE_FILE_PREFIX}_${docId}.json`
  const blob = Utilities.newBlob('{}', 'application/json', fileName)
  return DriveApp.createFile(blob)
}

function createTempDataFile_(docId) {
  const fileName = `${TEMP_DATA_PREFIX}_${docId}.json`
  const blob = Utilities.newBlob('{}', 'application/json', fileName)
  return DriveApp.createFile(blob)
}

function readState_(docId) {
  const file = getStateFile_(docId)
  if (!file) return null
  
  try {
    const content = file.getBlob().getDataAsString()
    const state = JSON.parse(content)
    
    // Load tagChildren from temp file if in writing phase
    if (state.phase === 'writing') {
      const tempFile = getTempDataFile_(docId)
      if (tempFile) {
        state.tagChildrenData = deserializeTagChildren_(tempFile)
      }
    }
    
    return state
  } catch (e) {
    Logger.log('Error reading state file: ' + e)
    return null
  }
}

function serializeElement_(element) {
  // Convert Document element to lightweight serializable format
  const elementType = element.getType()
  const typeString = elementType.toString()
  const serialized = {
    type: typeString
  }
  
  if (typeString === 'PARAGRAPH') {
    const para = element.asParagraph()
    serialized.text = para.getText()
    serialized.heading = para.getHeading().toString()
  } else if (typeString === 'LIST_ITEM') {
    const listItem = element.asListItem()
    serialized.text = listItem.getText()
    serialized.glyphType = listItem.getGlyphType().toString()
  } else if (typeString === 'INLINE_IMAGE') {
    // Store image as base64 to avoid slow Docs API calls
    const img = element.asInlineImage()
    serialized.blob = Utilities.base64Encode(img.getBlob().getBytes())
    serialized.width = img.getWidth()
    serialized.height = img.getHeight()
  } else {
    serialized.text = element.getText ? element.getText() : ''
  }
  
  return serialized
}

function deserializeElement_(serialized, doc) {
  // Reconstruct Document element from serialized format
  // Returns a simple object that can be used in the writing phase
  return {
    serializedType: serialized.type,
    text: serialized.text || '',
    heading: serialized.heading,
    glyphType: serialized.glyphType,
    imageBlob: serialized.blob,
    imageWidth: serialized.width,
    imageHeight: serialized.height
  }
}

function serializeTagChildren_(tagChildren, tempFile) {
  // Store tagChildren data as JSON (much faster than Docs API)
  const serializedData = {}
  
  Object.keys(tagChildren).forEach(tag => {
    serializedData[tag] = tagChildren[tag].map(tagChild => ({
      date: tagChild.date,
      elements: tagChild.elements.map(el => serializeElement_(el))
    }))
  })
  
  const content = JSON.stringify(serializedData)
  tempFile.setContent(content)
}

function deserializeTagChildren_(tempFile) {
  // Rebuild tagChildren structure from JSON file
  try {
    const content = tempFile.getBlob().getDataAsString()
    const serializedData = JSON.parse(content)
    const tagChildren = {}
    
    Object.keys(serializedData).forEach(tag => {
      tagChildren[tag] = serializedData[tag].map(tagChild => ({
        date: tagChild.date,
        elements: tagChild.elements.map(el => deserializeElement_(el))
      }))
    })
    
    return tagChildren
  } catch (e) {
    Logger.log('Error deserializing tagChildren: ' + e)
    return {}
  }
}

function writeState_(docId, state, tagChildren) {
  let file = getStateFile_(docId)
  if (!file) {
    file = createStateFile_(docId)
  }
  
  // If we have tagChildren to save, serialize them to temp file (fast JSON blob)
  if (tagChildren && Object.keys(tagChildren).length > 0) {
    let tempFile = getTempDataFile_(docId)
    if (!tempFile) {
      tempFile = createTempDataFile_(docId)
    }
    serializeTagChildren_(tagChildren, tempFile)
  }
  
  // Create a serializable version of the state (without tagChildren)
  const serializableState = {
    phase: state.phase,
    childIndex: state.childIndex || 0,
    childrenRemoved: state.childrenRemoved || 0,
    totalChildren: state.totalChildren || 0,
    lastDate: state.lastDate,
    inTagsSection: state.inTagsSection || false,
    tagsParagraphCreated: state.tagsParagraphCreated || false,
    currentTagIndex: state.currentTagIndex || 0,
    currentTagChildIndex: state.currentTagChildIndex || 0,
    sortedTags: state.sortedTags || []
  }
  
  const content = JSON.stringify(serializableState)
  file.setContent(content)
}

function deleteStateFiles_(docId) {
  const stateFile = getStateFile_(docId)
  if (stateFile) {
    stateFile.setTrashed(true)
  }
  
  const tempFile = getTempDataFile_(docId)
  if (tempFile) {
    tempFile.setTrashed(true)
  }
}

function shouldResumeFromState_(docId) {
  const file = getStateFile_(docId)
  if (!file) return false
  
  // Get document last modified date from Drive
  const docFile = DriveApp.getFileById(docId)
  const docLastModified = docFile.getLastUpdated()
  const stateLastModified = file.getLastUpdated()
  
  // Resume if state file is newer than the document
  return stateLastModified > docLastModified
}

function truncateText_(text, maxLength) {
  if (text.length <= maxLength) return text
  
  // Truncate at last word boundary before limit
  const truncated = text.substr(0, maxLength)
  const lastSpace = truncated.lastIndexOf(' ')
  const truncatedText = truncated.substr(0, lastSpace)
  const charsRemaining = text.length - truncatedText.length

  if (charsRemaining <= 0) return text
  
  return `${truncatedText + ELLIPSIS} [${charsRemaining} more characters...]`
}

function saveAndReopenDoc_(doc) {
  doc.saveAndClose()
  return DocumentApp.openById(doc.getId())
}

// Modified from https://fargyle.medium.com/apps-script-get-docs-heading-id-3ea2bde48778
function getHeadingId_(document, headingNamedRangeName) {
  var namedRangeArr;
  var headingId;
  var rangeStartIndex;
  var rangeEndIndex;
  var namedRanges = document.namedRanges;

  for (var namedRange in namedRanges) {
    if (namedRanges[namedRange].name == headingNamedRangeName) {
      namedRangeArr = namedRanges[namedRange].namedRanges;
      // There should only be one named range for this script, but a cleanup may have failed
      // So validate id; can't do beacuse Apps Script and Advanced API ids are different (see below)
      for (var i=namedRangeArr.length;i--;) {
        
        // In general there should only be one range per namedRange, but sometimes can be multiple
        // https://developers.google.com/docs/api/reference/rest/v1/NamedRange
        var rangeArr = namedRangeArr[i].ranges; 
        for (var j=rangeArr.length;j--;) { 
          // heading should be at first index, but just in case
          if (!rangeArr[j].segmentId) { // body
            rangeStartIndex = rangeArr[j].startIndex;
            rangeEndIndex = rangeArr[j].endIndex;
            break;
          }
        }
      }   
    }
  }
  
  if (rangeStartIndex && rangeEndIndex) { // should always be neither or both, but just in case...
    var structuralElement;
    var paragraphStyle;
    var contentArr = document.body.content;
    for (i=0;i<contentArr.length;i++) {
      structuralElement = contentArr[i];
      if (structuralElement.paragraph 
          && structuralElement.startIndex == rangeStartIndex) { // range starts right after the header
        paragraphStyle = structuralElement.paragraph.paragraphStyle;
        if (paragraphStyle) { // should always be true, but just in case...
          headingId = paragraphStyle.headingId;
        }
      }
    }
  }

  return headingId;
}

function findTagsAndBuildIndex() {
  let doc = DocumentApp.getActiveDocument()
  const docId = doc.getId()
  const startTime = Date.now()
  
  // Check if we should resume from a saved state
  let state = null
  if (shouldResumeFromState_(docId)) {
    state = readState_(docId)
    if (state) {
      Logger.log('Resuming from saved state: phase=' + state.phase)
    }
  }
  
  // Initialize state if starting fresh
  if (!state) {
    // Get total children count at the start
    const totalChildren = doc.getBody().getNumChildren()
    
    state = {
      phase: 'gathering',
      childIndex: 0,
      childrenRemoved: 0,
      totalChildren: totalChildren,  // Store original count
      tagChildren: {},
      currentTagMatches: [],
      lastDate: null,
      inTagsSection: false,
      tagsParagraphCreated: false,
      // Writing phase state
      sortedTags: [],
      currentTagIndex: 0,
      currentTagChildIndex: 0
    }
  }
  
  try {
    if (state.phase === 'gathering') {
      state = gatheringPhase_(doc, state, startTime, docId)
    }
    
    if (state.phase === 'writing') {
      // Restore tagChildren from temp doc if resuming
      if (state.tagChildrenData) {
        state.tagChildren = state.tagChildrenData
        state.sortedTags = Object.keys(state.tagChildren).sort()
      }
      
      state = writingPhase_(doc, state, startTime, docId)
    }
    
    // If we completed successfully, clean up the state files
    if (state.phase === 'complete') {
      deleteStateFiles_(docId)
      doc.saveAndClose()
      Logger.log('Indexing completed successfully')
    }
  } catch (e) {
    Logger.log('Error during indexing: ' + e)
    throw e
  }
}

function gatheringPhase_(doc, state, startTime, docId) {
  let changeCount = 0
  // Use the original totalChildren count from when we started
  const totalChildren = state.totalChildren
  
  for (var childIndex = state.childIndex; childIndex < totalChildren; childIndex++) {
    // Check if we're approaching the time limit
    // Only save state when not collecting multi-line tags to prevent data corruption
    // (currentTagMatches tracks in-progress multi-line tag collections)
    if (Date.now() - startTime > MAX_RUNTIME_MS && state.currentTagMatches.length === 0) {
      state.childIndex = childIndex
      writeState_(docId, state, state.tagChildren)
      doc.saveAndClose()
      Logger.log('Saved gathering phase state at child index: ' + childIndex)
      return state
    }
    
    const child = doc.getBody().getChild(childIndex - state.childrenRemoved)
    
    const isParagraph = child.getType() === DocumentApp.ElementType.PARAGRAPH
    if (isParagraph) {
      const paragraph = child.asParagraph()
      const pHeading = paragraph.getHeading()
      if (pHeading != null) {
        switch (pHeading) {
          case DocumentApp.ParagraphHeading.HEADING3:
            const heading = paragraph.getText()
            if (heading) {
              const ranges = doc.getNamedRanges(heading)
              if (ranges) {
                ranges.forEach(range => range.remove())
              }
              const rangeBuilder = doc.newRange();
              rangeBuilder.addElement(paragraph)
              doc.addNamedRange(heading, rangeBuilder.build())
              state.lastDate = heading
            }
            break
          case DocumentApp.ParagraphHeading.HEADING1:
            if (paragraph.getText() === TAGS_SECTION_NAME) {
              state.inTagsSection = true

              // append a new Tags section, it won't get deleted as we delete everything up to the last paragraph
              if (!state.tagsParagraphCreated) {
                doc.getBody().appendParagraph(TAGS_SECTION_NAME)
                state.tagsParagraphCreated = true
              }
            }
            break
        }
      }
    }
    
    if (!state.inTagsSection) {
      if (state.lastDate && child) {
        // Process multi-line tag matches
        // Iterate backwards to safely remove completed matches without index adjustment
        for (var tagMatchIndex = state.currentTagMatches.length - 1; tagMatchIndex >= 0; tagMatchIndex--) {
          const tagMatch = state.currentTagMatches[tagMatchIndex]
          tagMatch.elementDetails.elements.push(child.copy())
          const childrenRemaining = tagMatch.childrenRemaining - 1
          if (childrenRemaining === 0) {
            state.tagChildren[tagMatch.tag].push(tagMatch.elementDetails)
            state.currentTagMatches.splice(tagMatchIndex, 1)
          } else {
            tagMatch.childrenRemaining = childrenRemaining
          }
        }

        const tagMatches = child.getText().length && child.getText().match(TAGS_REGEX)
        if (tagMatches) {
          tagMatches.forEach(tagMatch => {
            const elementDetails = {date: state.lastDate, elements: [child.copy()]}
            const tagDetails = tagMatch.split('_')
            const tag = tagDetails[0]
            if (!state.tagChildren[tag]) {
              state.tagChildren[tag] = []
            }

            if (tagDetails[1] && tagDetails[1].startsWith('+')) {
              state.currentTagMatches.push({
                tag: tag,
                childrenRemaining: Number(tagDetails[1].substring(1)),
                elementDetails: elementDetails
              })
            } else {
              state.tagChildren[tag].push(elementDetails)
            }
          })
        }
      }
    } else {
      // if we've detected the Tags section, delete it and everything after so we can rebuild it later
      child.removeFromParent()
      state.childrenRemoved++
    }
  }
  
  // Gathering phase complete, move to writing phase
  state.phase = 'writing'
  state.sortedTags = Object.keys(state.tagChildren).sort()
  state.currentTagIndex = 0
  state.currentTagChildIndex = 0
  
  // Save state before transitioning to writing phase
  writeState_(docId, state, state.tagChildren)
  Logger.log('Gathering phase complete, transitioning to writing phase')
  
  return state
}

function writingPhase_(doc, state, startTime, docId) {
  const docsApiDoc = Docs.Documents.get(doc.getId())
  let changeCount = 0
  
  // Find or create the Tags header
  let tagsHeader = null
  const body = doc.getBody()
  const numChildren = body.getNumChildren()
  
  // Look for existing Tags header (should be at the end after gathering phase)
  for (let i = numChildren - 1; i >= 0; i--) {
    const child = body.getChild(i)
    if (child.getType() === DocumentApp.ElementType.PARAGRAPH) {
      const para = child.asParagraph()
      if (para.getHeading() === DocumentApp.ParagraphHeading.HEADING1 &&
          para.getText() === TAGS_SECTION_NAME) {
        tagsHeader = para
        break
      }
    }
  }
  
  if (!tagsHeader) {
    tagsHeader = body.appendParagraph(TAGS_SECTION_NAME)
    tagsHeader.setHeading(DocumentApp.ParagraphHeading.HEADING1)
  }
  
  // Process tags starting from where we left off
  for (let tagIndex = state.currentTagIndex; tagIndex < state.sortedTags.length; tagIndex++) {
    // Check if we're approaching the time limit
    if (Date.now() - startTime > MAX_RUNTIME_MS) {
      state.currentTagIndex = tagIndex
      writeState_(docId, state, null) // Don't re-save tagChildren, just state
      doc.saveAndClose()
      Logger.log('Saved writing phase state at tag index: ' + tagIndex)
      return state
    }
    
    const tag = state.sortedTags[tagIndex]
    const tagChildren = state.tagChildren[tag]
    
    // Add tag header if we're starting a new tag
    if (state.currentTagChildIndex === 0) {
      const tagHeader = body.appendParagraph(tag)
      tagHeader.setHeading(DocumentApp.ParagraphHeading.HEADING2)
      changeCount += 2
    }
    
    // Reverse the ordering so we get them in ascending order (earliest to most recent dates)
    const reversedChildren = tagChildren.slice().reverse()
    
    // Process tag children starting from where we left off
    for (let childIdx = state.currentTagChildIndex; childIdx < reversedChildren.length; childIdx++) {
      // Check if we're approaching the time limit
      if (Date.now() - startTime > MAX_RUNTIME_MS) {
        state.currentTagIndex = tagIndex
        state.currentTagChildIndex = childIdx
        writeState_(docId, state, null) // Don't re-save tagChildren, just state
        doc.saveAndClose()
        Logger.log('Saved writing phase state at tag: ' + tag + ', child: ' + childIdx)
        return state
      }
      
      const tagChild = reversedChildren[childIdx]
      
      if (changeCount >= SAVE_THRESHOLD) {
        doc.saveAndClose()
        doc = DocumentApp.openById(doc.getId())
        changeCount = 0
      }
      
      const p = body.appendParagraph('')
      const dateText = p.appendText(tagChild.date)

      // a new heading won't always show up on the first API call since the named range was just added
      const headingId = getHeadingId_(docsApiDoc, tagChild.date)
      if (headingId) dateText.setLinkUrl(`#heading=${headingId}`)

      dateText.setBold(true)
      tagChild.elements.forEach(child => {
        // Handle both regular Document elements and deserialized elements
        const isDeserialized = child.serializedType !== undefined
        
        if (isDeserialized) {
          // Element was loaded from state - use serialized data
          if (child.serializedType === 'INLINE_IMAGE') {
            if (child.imageBlob) {
              const imageBlob = Utilities.newBlob(Utilities.base64Decode(child.imageBlob))
              const image = body.appendImage(imageBlob)
              if (child.imageWidth) image.setWidth(child.imageWidth)
              if (child.imageHeight) image.setHeight(child.imageHeight)
            }
          } else {
            const truncated = truncateText_(child.text.replace(TAGS_REGEX, ''), MAX_CHARS_PER_ENTRY)
            if (child.serializedType === 'LIST_ITEM') {
              body.appendListItem(truncated)
            } else {
              body.appendParagraph(truncated)
            }
          }
        } else {
          // Element is from current gathering - use Document API
          switch (child.getType()) {
            case DocumentApp.ElementType.INLINE_IMAGE:
              body.appendImage(child.copy())
              break;
            case DocumentApp.ElementType.PARAGRAPH:
            case DocumentApp.ElementType.LIST_ITEM:
              const truncated = truncateText_(child.copy().getText().replace(TAGS_REGEX, ''), MAX_CHARS_PER_ENTRY)
              child.getType() === DocumentApp.ElementType.PARAGRAPH ?
                body.appendParagraph(truncated) :
                body.appendListItem(truncated)
              break;
          }
        }
      })
      body.appendParagraph('')
      changeCount += tagChild.elements.length + 2 // Approximate change count
    }
    
    // Reset child index for the next tag
    state.currentTagChildIndex = 0
  }
  
  // Writing phase complete
  state.phase = 'complete'
  return state
}