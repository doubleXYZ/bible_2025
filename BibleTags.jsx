// Преобразование тегов ссылок на сноски в форматированный текст
function BibleReferenceConverter() {

    var found, myStory, myStory, myRefIndex, myFirstIndex, myLastIndex;
    var myOutFirstIndex, myOutLastIndex;

    var myOpenTag = "<ref::>"
    var myCloseTag = "<::ref>"
	var myFindWhatStr = myOpenTag + "([\\*\\w]+?)" + myCloseTag
    var openTagLenth = myOpenTag.length; // 
    var closeTagLenth = myCloseTag.length; // 

    app.changeGrepPreferences = NothingEnum.nothing;
	app.findGrepPreferences = NothingEnum.nothing;
	app.findChangeGrepOptions.includeFootnotes = true;
	app.findChangeGrepOptions.includeHiddenLayers = false;
	app.findChangeGrepOptions.includeLockedLayersForFind = false;
	app.findChangeGrepOptions.includeLockedStoriesForFind = false;
	app.findChangeGrepOptions.includeMasterPages = false;

	app.findGrepPreferences.findWhat = myFindWhatStr;
	var myFoundItems = app.findGrep()

    // обработка найденного
    for (var i = myFoundItems.length - 1; i >= 0; i--) {
        found = myFoundItems[i];
        found.appliedCharacterStyle = "BibleReference";

        myStory = found.parentStory
		// myRefIndex = found.insertionPoints.item(0).index
		myOutFirstIndex = found.insertionPoints.item(0).index;
		myFirstIndex = found.insertionPoints.item(0).index + openTagLenth;
		myOutLastIndex = found.insertionPoints.item(-1).index;
		myLastIndex = found.insertionPoints.item(-1).index - closeTagLenth;
		myStory.insertionPoints.itemByRange(myFirstIndex, myLastIndex).select();
		app.copy();
		myStory.insertionPoints.itemByRange(myOutFirstIndex, myOutLastIndex).select();
        app.paste();
        
    }

}

BibleReferenceConverter()