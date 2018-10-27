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

var version_excel_gleaner_js = 3;

