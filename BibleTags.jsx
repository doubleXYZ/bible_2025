// Преобразование тегов ссылок на сноски в форматированный текст
// var mDoc = app.activeDocument

function main() {
  var doc = app.activeDocument;
  if (!doc) return;

  var storiesLength = doc.stories.length;
  // Обрабатываем все истории документа
  for (var s = 0; s < storiesLength; s++) {
    var story = doc.stories[s];
    BibleReferenceConverter(story);
  }
}

function BibleReferenceConverter(myStory) {
  var found, myRefIndex, myFirstIndex, myLastIndex;
  var myOutFirstIndex, myOutLastIndex;

  var myOpenRefTag = "<ref::>"; // начало ссылки на сноски
  var myCloseRefTag = "<::ref>"; // конец ссылки на сноски
  var myOpenTag = "<FootnoteStart>"; // начало тега сноски
  var myCloseTag = "</FootnoteEnd>"; // конец тега сноски
  var myFindWhatStr = myOpenTag + "([\\*\\w\\W]+?)" + myCloseTag;
  //   var openTagLenth = myOpenRefTag.length; //
  //   var closeTagLenth = myCloseRefTag.length; //

  var regex = /<ref::>(.+?)<::ref>/;

  resetFindGrep();
  app.findGrepPreferences.findWhat = myFindWhatStr;
  var myFoundItems = myStory.findGrep();

  // обработка найденного
  for (var i = myFoundItems.length - 1; i >= 0; i--) {
    found = myFoundItems[i];
    var refText = regex.exec(found.contents)[1];
    var refLength = refText.length;

    myStory = found.parentStory;
    myRefInsertionPoint = found.insertionPoints.lastItem();
    myRefInsertionPoint.contents = refText;
    myRefIndex = myRefInsertionPoint.index;
    myLastIndex = myRefInsertionPoint.index + refLength;

    var selectedRef;
    try {
      selectedRef = myStory.insertionPoints.itemByRange(
        myRefIndex,
        myLastIndex
      );
      selectedRef.appliedCharacterStyle =
        app.activeDocument.characterStyles.item("ddd");
    } catch (e) {
      alert(e);
    }
  }
}

function resetFindGrep() {
  app.changeGrepPreferences = NothingEnum.nothing;
  app.findGrepPreferences = NothingEnum.nothing;
  app.findChangeGrepOptions.includeFootnotes = true;
  app.findChangeGrepOptions.includeHiddenLayers = false;
  app.findChangeGrepOptions.includeLockedLayersForFind = false;
  app.findChangeGrepOptions.includeLockedStoriesForFind = false;
  app.findChangeGrepOptions.includeMasterPages = false;
}

// BibleReferenceConverter();

main();
