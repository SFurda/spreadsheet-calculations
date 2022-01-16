function getDirection() {

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var mapSheet = ss.getSheetByName("MAP1");


  for (var info = 2; info <= 39; info++) {

    var start = mapSheet.getRange('D1').getValue();
    var end = mapSheet.getRange(info, 2).getValue();
    var sites = mapSheet.getRange(info, 1).getValue();

    var directions = Maps.newDirectionFinder()
      .setOrigin(start)
      .setDestination(end)
      .setMode(Maps.DirectionFinder.Mode.DRIVING)
      .getDirections();

    // Logger.log(directions.routes[0].legs[0].duration.text);

   mapSheet.getRange('K1:N100').clear();

    var nextRow = mapSheet.getLastRow() + 1;

    for (var i = 0; i < directions.routes[0].legs.length; i++) {

      var endAddress = directions.routes[0].legs[i].end_address;
      var startAddress = directions.routes[0].legs[i].start_address;
      var distance = directions.routes[0].legs[i].distance.text;
      var duration = directions.routes[0].legs[i].duration.text;

      mapSheet.getRange(nextRow, 7).setValue(sites);
      mapSheet.getRange(nextRow, 8).setValue(endAddress);
      mapSheet.getRange(nextRow, 9).setValue(distance);
      mapSheet.getRange(nextRow, 10).setValue(duration);

    }


  }

}

