const TAGS_REGEX_STRING = "#[^\\s]+"
const TAGS_REGEX = new RegExp(TAGS_REGEX_STRING, 'g')
const TAGS_SECTION_NAME = "Tags"
const MAX_CHARS_PER_ENTRY = 150
const ELLIPSIS = "..."
const SAVE_THRESHOLD = 100

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
  let changeCount = 0
  var lastDate = null
  var currentTagMatches = []
  const tagChildren = {}
  var inTagsSection = false
  var tagsParagraph = null

  const totalChildren = doc.getBody().getNumChildren()
  var childrenRemoved = 0
  for (var childIndex = 0; childIndex < totalChildren; childIndex++) {
    // our childIndex keeps increasing, but the number of children in the document decreases if we remove them, 
    // (when inTagsSection is true) so the real index of the next child actually changes
    const child = doc.getBody().getChild(childIndex-childrenRemoved)
    
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
              lastDate = heading
            }
            break
          case DocumentApp.ParagraphHeading.HEADING1:
            if (paragraph.getText() === TAGS_SECTION_NAME) {
              inTagsSection = true

              // append a new Tags section, it won't get deleted as we delete everything up to the last paragraph
              tagsParagraph = doc.getBody().appendParagraph(TAGS_SECTION_NAME)
            }
            break
        }
      }
    }
    
    if (!inTagsSection) {
      if (lastDate && child) {
        for (var tagMatchIndex = 0; tagMatchIndex < currentTagMatches.length; tagMatchIndex++) {
          const tagMatch = currentTagMatches[tagMatchIndex]
          tagMatch.elementDetails.elements.push(child.copy())
          const childrenRemaining = tagMatch.childrenRemaining-1
          if (childrenRemaining === 0) {
            tagChildren[tagMatch.tag].push(tagMatch.elementDetails)
            currentTagMatches.splice(tagMatchIndex, 1)
          } else {
            tagMatch.childrenRemaining = childrenRemaining
          }
        }

        const tagMatches = child.getText().length && child.getText().match(TAGS_REGEX)
        if (tagMatches) {
          tagMatches.forEach(tagMatch => {
            const elementDetails = {date: lastDate, elements: [child.copy()]}
            const tagDetails = tagMatch.split('_')
            const tag = tagDetails[0]
            if (!tagChildren[tag]) {
              tagChildren[tag] = []
            }

            if (tagDetails[1] && tagDetails[1].startsWith('+')) {
              currentTagMatches.push({
                tag: tag,
                childrenRemaining: Number(tagDetails[1].substring(1)),
                elementDetails: elementDetails
              })
            } else {
              tagChildren[tag].push(elementDetails)
            }
          })
        }
      }
    } else {
      // if we've detected the Tags section, delete it and everything after so we can rebuild it later
      child.removeFromParent()
      childrenRemoved++
    }
  }

  const docsApiDoc = Docs.Documents.get(doc.getId())
  const tagsHeader = tagsParagraph || doc.getBody().appendParagraph(TAGS_SECTION_NAME)
  tagsHeader.setHeading(DocumentApp.ParagraphHeading.HEADING1)
  Object.keys(tagChildren).sort().forEach(tag => {
    const tagHeader = doc.getBody().appendParagraph(tag)
    tagHeader.setHeading(DocumentApp.ParagraphHeading.HEADING2)
    changeCount += 2

    // reverse the ordering so we get them in ascending order (earliest to most recent dates)
    tagChildren[tag].reverse().forEach(tagChild => {
      if (changeCount >= SAVE_THRESHOLD) {
        doc = saveAndReopenDoc_(doc)
        changeCount = 0
      }
      const p = doc.getBody().appendParagraph('')
      const dateText = p.appendText(tagChild.date)

      // a new heading won't always show up on the first API call since the named range was just added
      const headingId = getHeadingId_(docsApiDoc, tagChild.date)
      if (headingId) dateText.setLinkUrl(`#heading=${headingId}`)

      dateText.setBold(true)
      tagChild.elements.forEach(child => {
        switch (child.getType()) {
          case DocumentApp.ElementType.INLINE_IMAGE:
            doc.getBody().appendImage(child.copy())
            break;
          case DocumentApp.ElementType.PARAGRAPH:
          case DocumentApp.ElementType.LIST_ITEM:
            const truncated = truncateText_(child.copy().getText().replace(TAGS_REGEX, ''), MAX_CHARS_PER_ENTRY)
            child.getType() === DocumentApp.ElementType.PARAGRAPH ?
              doc.getBody().appendParagraph(truncated) :
              doc.getBody().appendListItem(truncated)
            break;
        }
      })
      doc.getBody().appendParagraph('')
      changeCount += tagChild.elements.length + 2 // Approximate change count
    })
  })
  doc.saveAndClose()
}