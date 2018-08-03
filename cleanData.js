<script type="text/javascript">
  var ws_data = [];
  var farmerNameExtr = $("placemark name");
  var farmerCoorExtr = $("placemark Polygon  coordinates");
  farmerNameExtr.splice(74, 1);
  farmerCoorExtr.splice(74, 2);
  farmerNameExtr.splice(127, 1);
  farmerCoorExtr.splice(127, 2);

  console.log("Farmers name number: " + farmerNameExtr);
  console.log("Farms coordinates number: " + farmerCoorExtr);

  for(var i=0; i<farmerCoorExtr.length; i++){
    var reSendData = Array();
    var coorIniti = farmerCoorExtr[i].textContent;
    coorSpli = coorIniti.split(" ");
    revNull = coorSpli.filter(String);
    revOffset = revNull;
    revOffset.splice(0,1);
    for(var j=0; j< revOffset.length; j++){
      String(revOffset[j]).slice(0, revOffset[j].length-3);
      coorSpli_2 = revOffset[j].split(",");
      mergeSet = "[" + coorSpli_2[1] + "," + coorSpli_2[0] + "]";
      reSendData.push(mergeSet);
      joinData = reSendData.join();
    }
    joinData = "[" + joinData + "]";
    //console.log(farmerNameExtr[i].textContent);
    //console.log(joinData);

    //Write Excel
    FinalData = [farmerNameExtr[i].textContent, joinData];
    ws_data.push(FinalData);

  }
  var wb = XLSX.utils.book_new();
  wb.Props = {
          Title: "SheetJS Tutorial",
          Subject: "Test",
          Author: "Red Stapler",
          CreatedDate: new Date(2017,12,19)
  };

  wb.SheetNames.push("Test Sheet");

  var ws = XLSX.utils.aoa_to_sheet(ws_data);
  wb.Sheets["Test Sheet"] = ws;

  var wbout = XLSX.write(wb, {bookType:'xlsx',  type: 'binary'});
  function s2ab(s) {

          var buf = new ArrayBuffer(s.length);
          var view = new Uint8Array(buf);
          for (var i=0; i<s.length; i++) view[i] = s.charCodeAt(i) & 0xFF;
          return buf;

  }
  $("#button-a").click(function(){
          saveAs(new Blob([s2ab(wbout)],{type:"application/octet-stream"}), 'finalCoordiData.xlsx');
  });
  //Learn how to write excel by javascript.
  //Cut outlier out.
</script>
