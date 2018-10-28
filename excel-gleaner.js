
var a = null;

function downloadFile(data, fileName) {
    a = document.createElement('a');
    document.body.appendChild(a);
    a.download = fileName;
    var encodedData = btoa(unescape(encodeURIComponent(data)));
    a.href = "data:application/octet-stream;charset=utf-8;base64,"+encodedData;
    a.click();   

   //data:application/octet-stream;charset=utf-8;base64,Zm9vIGJ
}

//
function readBlob() {
    var files = document.getElementById('files').files;
    if (!files.length) {
        document.getElementById('msg').textContent = 'Please select a file!';
        return;
    }

    var file = files[0];
    var start =  0;
    var stop = file.size - 1;

    var reader = new FileReader();

    // If we use onloadend, we need to check the readyState.
    reader.onloadend = function(evt) {
        if (evt.target.readyState == FileReader.DONE) { // DONE == 2
        // document.getElementById('byte_content').textContent = evt.target.result;
        document.getElementById('byte_range').textContent = 
            ['Read bytes: ', start + 1, ' - ', stop + 1,
                ' of ', file.size, ' byte file'].join('');
                document.getElementById('msg').textContent = "... reading done."
        }
    };
    var blob = file.slice(start, stop + 1);
    reader.readAsBinaryString(blob);
} // readBlob()

document.querySelector('.readBytesButtons').addEventListener('click', function(evt) {
    if (evt.target.tagName.toLowerCase() == 'button') {
        readBlob();
    }
}, false);

function showSheets()
{
    var files = document.getElementById('files').files;
    if (!files.length) {
        document.getElementById('msg').textContent = 'Please select a file!';
        return;
    }
    var file = files[0]; 
    var reader = new FileReader();
    reader.onload = function(e) {
        var data = e.target.result; // not working in chrome: result();
        var workbook = XLSX.read(data, {
          type: 'binary'
        });
        var sheets = "";
        workbook.SheetNames.forEach(function(sheetName) {
          // Here is your object
          //var XL_row_object = XLSX.utils.sheet_to_row_object_array(workbook.Sheets[sheetName]);
          //var json_object = JSON.stringify(XL_row_object);
          //console.log(json_object);
          //console.log("sheetName='"+sheetName+"'");
          sheets += ""+sheetName+", ";
        })
        document.getElementById('msg').textContent = "found sheets: " + sheets;
  
      }; // reader.onload
  
      reader.onerror = function(ex) {
        document.getElementById('msg').textContent = ""+ex+"";
        //console.log(ex);
      };
  
      reader.readAsBinaryString(file);
} // showSheets()


document.querySelector('.excel').addEventListener('click', function(evt) {
    if (evt.target.tagName.toLowerCase() == 'button') {
        showSheets();
    }
}, false);


var debugValue = null;
function exportFirstSheet()
{
    var files = document.getElementById('files').files;
    if (!files.length) {
        document.getElementById('msg').textContent = 'Please select a file!';
        return;
    }
    var file = files[0]; 
    var reader = new FileReader();
    reader.onload = function(e) {
        var data = e.target.result; // not working in chrome: result();
        var workbook = XLSX.read(data, {
          type: 'binary'
        });
        var sheets = "";
        var sheetNumber = 1;
        workbook.SheetNames.forEach(function(sheetName) {
          if (sheetNumber == 1)
          {
            // export
            var XL_row_object = XLSX.utils.sheet_to_row_object_array(workbook.Sheets[sheetName]);
           // debugValue = XL_row_object;   
            var json_object = JSON.stringify(XL_row_object, null, 2);
            //debugValue = json_object;
            document.getElementById('msg').textContent = "Excel file parsed. Found "+XL_row_object.length + 
                " rows in the first sheet.";
            var javascriptSource = "var "+sheetName+" = \r\n"+
               json_object.replace(new RegExp("\n", 'g'),"\r\n") + "; \r\n";
            downloadFile(javascriptSource, sheetName+".json.txt");
          }
          sheetNumber += 1;
        })

      }; // reader.onload
  
      reader.onerror = function(ex) {
        document.getElementById('msg').textContent = ""+ex+"";
        //console.log(ex);
      };
  
      reader.readAsBinaryString(file);
} // exportFirstSheet()



document.querySelector('.excel2').addEventListener('click', function(evt) {
    if (evt.target.tagName.toLowerCase() == 'button') {
        exportFirstSheet();
    }
}, false);


var version_excel_gleaner_js = 3;

