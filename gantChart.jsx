#target illustrator

var docRef = app.activeDocument;
var artboard = docRef.artboards[docRef.artboards.getActiveArtboardIndex()];

var startDate = new Date(2018, 9, 1) ;
var endDate = new Date(2019, 2, 1);

var months = ["JAN","FEB","MAR","APR","MAY","JUN","JUL","AUG","SEP","OCT","NOV","DEC"];

// Set Basic Colors
var gray = new RGBColor();
gray.red = 50;
gray.blue = 50;
gray.green = 50;
var white = new RGBColor();
white.red = 255;
white.blue = 255;
white.green = 255;

var chart = new Chart(startDate, endDate, artboard);
chart.drawGrid();
chart.drawMonthHeader();


function Chart(startDate, endDate, artboard) {
  this.left = artboard.artboardRect[0] + 36;
  this.right = artboard.artboardRect[2] - 36;
  this.top = artboard.artboardRect[1] - 50;
  this.width = this.right - this.left;
  this.headerOffset = 36;
  
  this.totalDays = Math.floor((endDate.getTime() - startDate.getTime()) / 86400000);
  this.rectWidth = this.width / this.totalDays;

  this.drawMonthHeader = function() {
    var currentX = this.left;
    var currentY = this.top;
    var rectHeight = 36;

    var totalDays = this.totalDays;
    var current = startDate;
    var targetLayer = docRef.layers.add();
    targetLayer.name = "Month Header";

    while (totalDays > 0) {
      var cM, cY, cD, eM, eY, eD;
      cM = current.getMonth() + 1;
      cY = current.getFullYear();
      cD = current.getDate();
      eM = endDate.getMonth();
      eY = endDate.getFullYear() + 1;
      eD = endDate.getDate();

      if (cM == eM && cY == eY) {
          var offset = endDate;
      } else {
          var m = (cM == 12) ? 1 : cM + 1;
          var y = (cM == 12) ? cY + 1 : cY; 
          var offset = new Date(y, m - 1, 1);
      }

      var daysUntilNext = Math.floor((offset.getTime() - current.getTime()) / 86400000);
      var w = daysUntilNext * this.rectWidth;

      var headerRect = targetLayer.pathItems.rectangle(currentY, currentX, w, rectHeight);
      headerRect.stroked = true;
      headerRect.filled = true;
      headerRect.fillColor = gray;
      headerRect.strokeColor = white;
      var rectRef = targetLayer.pathItems.rectangle(currentY - rectHeight * 0.4, currentX, w, rectHeight);
      var areaTextRef = targetLayer.textFrames.areaText(rectRef);
      areaTextRef.paragraphs.add(months[current.getMonth()]);
      applyAttributes(areaTextRef);

      current = offset;
      totalDays -= daysUntilNext;
      currentX += w;
    }
  }

  this.drawGrid = function() {
    var currentX = this.left;
    var top = this.top - (this.headerOffset + 10);
    var targetLayer = docRef.layers.add();
    targetLayer.name = "GridLines";
    for(var i = 0; i < this.totalDays + 1; i++) {
      var lineList = [
        [currentX, top],
        [currentX, artboard.artboardRect[3] + 36],
      ];
      var dayPath = targetLayer.pathItems.add();
      dayPath.setEntirePath(lineList);
      currentX += this.rectWidth;
    }
  }
}

function applyAttributes(areaTextRef) {
  var fontStyle = areaTextRef.textRange.characterAttributes;
  var pStyle = areaTextRef.textRange.paragraphs[0].paragraphAttributes;
  fontStyle.textFont = app.textFonts.getByName("Helvetica-Bold");
  fontStyle.capitalization = FontCapsOption.ALLCAPS;
  fontStyle.size = 18;
  fontStyle.fillColor = white;
  pStyle.justification = Justification.CENTER;
}

