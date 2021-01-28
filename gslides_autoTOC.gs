// adds menu item and function to create a TOC slide

function onOpen() {
  var ui = SlidesApp.getUi();
  ui.createMenu("TOC").addItem("Insert TOC slide", "toc").addToUi();
}

function toc() {
  var lineHeight = 25;
  var lineWidth = 1000;
  var margin = 50;
  var yPos = 100; //intitial value to place first text box vertically

  var slides = SlidesApp.getActivePresentation();
  var ui = SlidesApp.getUi();
  var theSlides = slides.getSlides();
  var len = theSlides.length;

  var toc = slides.insertSlide(0);
  toc.getShapes()[0].getText().setText("TOC");

  for (var i = 1; i < len; i++) {
    var j = i + 1; //slide number
    var yPos = yPos + lineHeight;
    var slide = theSlides[i];
    try {
      var title = slide.getShapes()[0].getText().asString();
      if (title == "") title = j;
      toc
        .insertShape(
          SlidesApp.ShapeType.TEXT_BOX,
          margin,
          yPos,
          lineWidth,
          lineHeight
        )
        .getText()
        .setText(title);
      toc.getShapes()[i].getText().getTextStyle().setLinkSlide(slide);
    } catch (err) {
      //if there is no title
      var title = "slide " + j;
    }
  }
}