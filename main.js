// javascript file. Load !!after!! jquery and fabric

// pivotelements = tabs in add-in
var PivotElements = document.querySelectorAll(".ms-Pivot");
for (var i = 0; i < PivotElements.length; i++) {
  new fabric['Pivot'](PivotElements[i]);
}

// Coverpage generator
// =============================

// dropdown vakken
var DropdownHTMLElements = document.querySelectorAll('.ms-Dropdown');

for (var i = 0; i < DropdownHTMLElements.length; ++i) {
  var Dropdown = new fabric['Dropdown'](DropdownHTMLElements[i]);
};

// hero button
var ButtonElements = document.querySelectorAll(".ms-Button");
  for (var i = 0; i < ButtonElements.length; i++) {
    new fabric['Button'](ButtonElements[i], function(){replaceContentInControl()});
  };

// datepicker
var DatePickerElements = document.querySelectorAll(".ms-DatePicker");
  for (var i = 0; i < DatePickerElements.length; i++) {
    new fabric['DatePicker'](DatePickerElements[i]);
  };


var replaceContentInControl = function() {
    Word.run(function (context) {
      var coursename = $('#cp-vak .ms-Dropdown-title').first().text();
      const serviceNameContentControl = context.document.contentControls.getByTag("cc_coursename").getFirst();
      serviceNameContentControl.insertText(coursename, "Replace");

        return context.sync();
    })
    .catch(function (error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    });
};

