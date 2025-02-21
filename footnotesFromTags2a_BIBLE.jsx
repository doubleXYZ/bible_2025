// footnotesFromTags.jsx
//An InDesign CS4 JavaScript
//
/*
Скрипт предназначен для восстановления сносок, помеченных в тексте тэгами.
Особенность скрипта в том, что он сохраняет локальное форматирование сносок Bold, Italic, Bold Italic и др.
Скрипт подразумевает, что преобразование сносок в тэгированный текст 
выполнено в MS Word макросом footnotesToText .
Mакрос для MS Word приведен в конце этого скрипта 
Скрипт пропускает сноски, находящиеся в таблицах.
Для запуска скрипта поставьте курсор в текст или выделите черной стрелкой любой текстовый фрейм Story.
(с) Борис Кащеев, www.adobeindesign.ru, boriskasmoscow@gmail.com
Благодарности:
Юрию Васильева из г. Киев за большой вклад в преодоление ошибок.
Дмитрию Глазкову г. Москва, за идею и тестирование
Александру Цветкову г. Москва, за участие в тестировании.
*/
// отключаем обновление экрана на время выполнения скрипта
app.scriptPreferences.enableRedraw = false;

// важные переменные
var bibleCharTagName = "BibleReference"; // имя стиля символа ссылки на сноску
var bibleCharTagName_inText = "BibleReference_inText"; // имя стиля символа ссылки на сноску

// ---------------------------------------------
// Задаем описание тэгов открытия/закрытия,
// которые используются в InDesign-документе
// ---------------------------------------------
var myOpenTag = "<FootnoteStart>";
var myCloseTag = "</FootnoteEnd>";

Object.prototype.isText = function () {
  switch (this.constructor.name) {
    case "InsertionPoint":
    case "Character":
    case "Word":
    case "TextStyleRange":
    case "Line":
    case "Paragraph":
    case "TextColumn":
    case "Text":
    case "TextFrame":
      return true;
    default:
      return false;
  }
}; // Object.prototype

function checkCharStyles(mDoc) {
  // ----- создание стилей символов 'bold', 'italic', 'bold italic' ------
  // ------------   если отсутствуют в документе   --------
  var characterStyleList = ["bold", "italic", "bold italic"];
  var fontStyleNames = ["Bold", "Italic", "Bold Italic"];
  for (var indx = 0; indx < characterStyleList.length; indx++) {
    var myCharacterStyleName = characterStyleList[indx];
    var myCharacterStyle;
    // Create a character style named "myCharacterStyleName" if
    // no style by that name already exists.
    if (!mDoc.characterStyles.item(myCharacterStyleName).isValid) {
      // If the character style does not exist, trying to get its name will generate an  error.
      myCharacterStyle = mDoc.characterStyles.add({
        name: myCharacterStyleName,
        fontStyle: fontStyleNames[indx],
      });
    }
  }
  // ----- создание стилей символов 'superscript' и 'subscript' ------
  // ------------   если отсутствуют в документе   --------

  var positionStyle = ["superscript", "subscript"];
  for (indx = 0; indx < 2; indx++) {
    if (!mDoc.characterStyles.item(positionStyle[indx]).isValid) {
      var myCharacterPositionStyle = mDoc.characterStyles.add({
        name: positionStyle[indx],
      });
      var myChangingCharStyle = mDoc.characterStyles.item(positionStyle[indx]);
      myChangingCharStyle.position = Position[positionStyle[indx]];
    }
  }

  // --------- проверка и создание в случае отсутствия
  // --------- стиля для библейской сноски BibleReference -----
  if (!mDoc.characterStyles.item("BibleReference").isValid) {
    // If the character style does not exist, trying to get its name will generate an  error.
    var bibleRefColor = mDoc.colors.add({
      name: "bibleRefColor",
      colorValue: [70, 0, 100, 0],
      model: ColorModel.PROCESS,
      space: ColorSpace.CMYK,
    });
    myCharacterStyle = mDoc.characterStyles.add({
      name: bibleCharTagName,
      appliedFont: "Octava",
      fontStyle: "Regular",
      pointSize: 8,
      baselineShift: 2,
      fillColor: bibleRefColor,
    });
  }
  if (!mDoc.characterStyles.item("BibleReference_inText").isValid) {
    // If the character style does not exist, trying to get its name will generate an  error.
    var bibleRefColor_inText = mDoc.colors.add({
      name: "bibleRefColor_inText",
      colorValue: [0, 100, 100, 0],
      model: ColorModel.PROCESS,
      space: ColorSpace.CMYK,
    });
    myCharacterStyle = mDoc.characterStyles.add({
      name: bibleCharTagName_inText,
      appliedFont: "Octava",
      fontStyle: "Regular",
      pointSize: 10,
      baselineShift: 3,
      fillColor: bibleRefColor_inText,
    });
  }
}

function findAndChangeCharaterStyle(mDoc) {
  resetFindTextPref();

  var charStylesList = ["italic", "bold", "bold italic"];
  var fontStyleNames = ["Italic", "Bold", "Bold Italic"];
  var property = "fontStyle";
  for (var charIdx = 0; charIdx < charStylesList.length; charIdx++) {
    var charStyle = mDoc.characterStyles.item(charStylesList[charIdx]);

    // Search for italic text in the selected text
    var foundItems = findChange("fontStyle", fontStyleNames[charIdx], mDoc); //app.findText();

    // Display the results
    if (foundItems.length > 0) {
      for (var i = 0; i < foundItems.length; i++) {
        foundItems[i].appliedCharacterStyle = charStyle;
      }
    } else {
      myMessage += 'No "' + fontStyleNames[charIdx] + '"\n';
    }

    // Reset find/change preferences
    resetFindTextPref();
  }
}

function findPositionStyle(mDoc) {
  resetFindTextPref();

  var positionStyle = ["superscript", "subscript"];
  var valueOfProperty = [1936749411, 1935831907]; // [Position.SUPERSCRIPT, Position.SUBSCRIPT]

  for (var i = 0; i < positionStyle.length; i++) {
    var charStyle = mDoc.characterStyles.item(positionStyle[i]);
    if (!charStyle) {
      alert(charStyle);
      return;
    }

    var foundItems = findChange("position", valueOfProperty[i], mDoc); // app.findText();
    if (foundItems.length > 0) {
      for (var idx = 0; idx < foundItems.length; idx++) {
        foundItems[idx].appliedCharacterStyle = charStyle;
      }
    } else {
      myMessage += 'No "' + positionStyle[i] + "\n";
    }
    resetFindTextPref();
  }
}

function findChange(property, valueOfProperty, curDoc) {
  // alert(property + ' - ' + valueOfProperty)
  var fCh = app.findTextPreferences;
  fCh[property] = valueOfProperty;
  var foundItems = curDoc.findText();

  return foundItems;
}

function resetFindTextPref() {
  app.changeTextPreferences.changeTo = "";
  app.findTextPreferences.findWhat = "";
  app.findTextPreferences = NothingEnum.nothing;
  app.changeTextPreferences = NothingEnum.nothing;

  app.findChangeTextOptions.includeFootnotes = true;
  app.findChangeTextOptions.includeHiddenLayers = false;
  app.findChangeTextOptions.includeLockedLayersForFind = false;
  app.findChangeTextOptions.includeLockedStoriesForFind = false;
  app.findChangeTextOptions.includeMasterPages = false;

  app.findChangeTextOptions.caseSensitive = false;
  app.findChangeTextOptions.wholeWord = false;
}
function showMyMessage() {
  if (myMessage.length > 0) {
    myMessage += " text found in selected text.";
    alert(myMessage);
  }
}

// ---------------------------------------------

//const myOpenTag = "<@F "
//const myCloseTag = ">"

main();
function main() {
  if (app.selection.length > 0 && app.selection[0].isText()) {
    var myDoc = app.documents[0];
    var myStory = app.selection[0].parentStory;

    // проверяем есть ли символьные стили в документе; если нет, то генерируем их
    checkCharStyles(myDoc);
    BibleReferenceConverter(myStory);
    footnotesFromTags(myDoc);
  } else {
    alert("Должен быть выделен текстовый объект");
  }
} // main()

function BibleReferenceConverter(myStory) {
  var found, myRefFirstIndex, myFirstIndex, myRefLastIndex;
  var myOutFirstIndex, myOutLastIndex;

  // var myOpenRefTag = "<ref::>"; // начало ссылки на сноски
  // var myCloseRefTag = "<::ref>"; // конец ссылки на сноски
  // var myOpenTag = "<FootnoteStart>"; // начало тега сноски
  // var myCloseTag = "</FootnoteEnd>"; // конец тега сноски
  var myFindWhatStr = myOpenTag + "([\\*\\w\\W]+?)" + myCloseTag;
  //   var openTagLenth = myOpenRefTag.length; //
  //   var closeTagLenth = myCloseRefTag.length; //

  var regex = /<ref::>(.+?)<::ref>/;

  resetGrepFind();
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
    myRefFirstIndex = myRefInsertionPoint.index;
    myRefLastIndex = myRefInsertionPoint.index + refLength;

    var selectedRef;
    try {
      selectedRef = myStory.insertionPoints.itemByRange(
        myRefFirstIndex,
        myRefLastIndex
      );
      selectedRef.appliedCharacterStyle =
        app.activeDocument.characterStyles.item("BibleReference_inText");
    } catch (e) {
      alert(e);
    }

    refChange(myStory, found, refLength);
  } // for()
  resetGrepFind();
} // BibleReferenceConverter()

function refChange(myStory, selectedText, refLength) {
  var myOpenRefTag = "<ref::>"; // начало ссылки на сноски
  var myCloseRefTag = "<::ref>"; // конец ссылки на сноски
  //    try {

  var myRefInsertionPoint =
    selectedText.insertionPoints.firstItem().index +
    myOpenTag.length +
    myOpenRefTag.length;
  var myRefEndInsertionPoint = myRefInsertionPoint + refLength;
  var selectedRef = myStory.insertionPoints.itemByRange(
    myRefInsertionPoint,
    myRefEndInsertionPoint
  );
  var charStyle = app.activeDocument.characterStyles.item("BibleReference");
  selectedRef.appliedCharacterStyle = charStyle;

  /*  
      } catch (e) {
        alert(e);
      }
    // */
  var findWhat = myOpenRefTag;
  resetFindTextPref();
  app.findTextPreferences.findWhat = findWhat;
  app.changeTextPreferences.changeTo = "";
  selectedText.changeText();

  var findWhat = myCloseRefTag;
  resetFindTextPref();
  app.findTextPreferences.findWhat = findWhat;
  app.changeTextPreferences.changeTo = "\t";
  myCloseRef = selectedText.findText();
  //   alert(myCloseRef)
  selectedText.changeText();
}

function footnotesFromTags(myDocument) {
  //------------количество символов в тэгах---------------------------------------
  var mySizeOfOpenTag = myOpenTag.length;
  var mySizeOfCloseTag = myCloseTag.length;
  //---------строка для поиска тэгов в тексте-----------------------------
  //var myFindWhatStr = myOpenTag + "([^<]+)" + myCloseTag
  var myFindWhatStr = myOpenTag + "([\\w\\W]+?)" + myCloseTag;
  // --------поиск коллекции тэгированных сносок -------------
  resetGrepFind();
  app.findGrepPreferences.findWhat = myFindWhatStr;
  var myFoundItems = app.findGrep();
  // --------данные для прогрессбара --------------
  var myProgressBarWidth = 300;
  var myMaximumValue = myFoundItems.length;
  var myIncrement = myProgressBarWidth / myMaximumValue;
  var myProgressPanel = myCreateProgressPanel(myProgressBarWidth);
  myProgressPanel.show();
  myProgressPanel.myProgressBar.value = myProgressBarWidth;
  myProgressPanel.currentFootnote.enabled = false;
  myProgressPanel.currentFootnote.text = "Restore Footnotes";

  // --------процесс преобразования сносок и вставки их в текст --------
  var myStory, myInsertionPoint;
  var myFirstIndex, myRefIndex, myLastIndex;
  var myFootnote;
  var counter = 0;
  for (var i = myFoundItems.length - 1; i >= 0; i--) {
    myProgressPanel.currentFootnote.text =
      "Now processing footnote: " + (i + 1) + " from " + myFoundItems.length;
    found = myFoundItems[i];
    if (found.parent.constructor.name == "Cell") continue; // пропускаем сноски в таблицах

    myStory = found.parentStory;
    myRefIndex = found.insertionPoints.item(0).index;
    myFirstIndex = found.insertionPoints.item(0).index + mySizeOfOpenTag;
    myLastIndex = found.insertionPoints.item(-1).index - mySizeOfCloseTag;
    myStory.insertionPoints.itemByRange(myFirstIndex, myLastIndex).select();
    app.copy();
    found.contents = "";
    myInsertionPoint = myStory.insertionPoints.item(myRefIndex);
    myFootnote = myInsertionPoint.footnotes.add();
    myFootnote.insertionPoints.item(-1).select();
    app.paste();
    myProgressPanel.myProgressBar.value = myIncrement * i;
    counter++;
  } // for
  myProgressPanel.hide();
  // ---------------Результаты работы------------------
  if (counter == myFoundItems.length) {
    alert(
      "Все сноски успешно восстановлены, " + counter + " шт.",
      "Поздравляем!"
    );
  } else {
    alert(
      "Восстановлено сносок " +
        counter +
        " из " +
        myFoundItems.length +
        "\rНевосстановленные сноски вероятно находятся в таблицах "
    );
  }
  // Конец
} // footnotesFromTags()

function resetGrepFind() {
  app.changeGrepPreferences = NothingEnum.nothing;
  app.findGrepPreferences = NothingEnum.nothing;
  app.findChangeGrepOptions.includeFootnotes = false;
  app.findChangeGrepOptions.includeHiddenLayers = false;
  app.findChangeGrepOptions.includeLockedLayersForFind = false;
  app.findChangeGrepOptions.includeLockedStoriesForFind = false;
  app.findChangeGrepOptions.includeMasterPages = false;
}

function myCreateProgressPanel(myProgressBarWidth) {
  myProgressPanel = new Window("window", "Progress Converting");
  with (myProgressPanel) {
    myProgressPanel.myProgressBar = add(
      "progressbar",
      [12, 12, myProgressBarWidth, 24],
      0,
      myProgressBarWidth
    );
    myProgressPanel.currentFootnote = add(
      "edittext",
      [12, 12, myProgressBarWidth, 36],
      ""
    );
  }
  return myProgressPanel;
} // fn

function quoteReplace(myDoc) {
  var replaceList = [
    {
      find: "\\[«\\]",
      replace: "„",
    },
    {
      find: "\\[/»\\]",
      replace: "“",
    },
    {
        find: "\\[„\\]",
        replace: "„",
      },
      {
        find: "\\[/“\\]",
        replace: "“",
      },
  ];

  for (var i = 0; i < replaceList.length; i++) {
    var find = replaceList[i].find;
    var replace = replaceList[i].replace;

    resetGrepFind();
    app.findGrepPreferences.findWhat = find;
    app.changeGrepPreferences.changeTo = replace;
    myDoc.changeGrep();
  }
}

function mainStyles() {
  if (app && app.name === "Adobe InDesign") {
    // // ----------- установка едениц измерения линеек --------------
    // app.documents[0].viewPreferences.horizontalMeasurementUnits = MeasurementUnits.MILLIMETERS;
    // app.documents[0].viewPreferences.verticalMeasurementUnits = MeasurementUnits.MILLIMETERS;

    if (app.selection.length > 0 && app.selection[0].isText()) {
      //~             var doc = app.documents[0];
      var doc = app.activeDocument;

      // ищем и меняем italic, bold, bold italic
      findAndChangeCharaterStyle(doc);
      // ищем и меняем superscript, subscript
      findPositionStyle(doc);
      quoteReplace(doc);
      showMyMessage();
    } else {
      alert("Должен быть выделен текстовый объект");
    }
  } else {
    alert("Adobe InDesign is not running.");
  }
}
var myMessage = "";
mainStyles();

// включаем обновление экрана на время выполнения скрипта
app.scriptPreferences.enableRedraw = true;

/*
Макрос для MS Word по преобразованию сносок в текст с сохранением локального форматирования (эта строка не принадлежит макросу)

Sub FootnotesToText()
' footnotes to text
Dim actdoc As Document
Dim fn As Word.Footnote
Dim rngFN As Word.Range
Dim i As Long
Set actdoc = ActiveDocument

For i = actdoc.Footnotes.Count To 1 Step -1
  Set fn = actdoc.Footnotes(i)  '
  Set rngFN = fn.Reference  '
  rngFN.Collapse wdCollapseEnd  '
  rngFN.FormattedText = fn.Range.FormattedText
  rngFN.InsertBefore Text:="<FootnoteStart:>"  
  rngFN.InsertAfter Text:="<FootnoteEnd:>"
  fn.Delete '
Next i
End Sub	
*/
