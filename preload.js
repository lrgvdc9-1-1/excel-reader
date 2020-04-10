const {shell, dialog } = require('electron').remote
const unzipper = require('unzipper');
const fs = require('fs');

function sleep(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

const removeDir = function(path) {
  if (fs.existsSync(path)) {
    const files = fs.readdirSync(path)
 
    if (files.length > 0) {
      files.forEach(function(filename) {
        if (fs.statSync(path + "/" + filename).isDirectory()) {
          removeDir(path + "/" + filename)
        } else {
          fs.unlinkSync(path + "/" + filename)
        }
      })
      fs.rmdirSync(path)
    } else {
      fs.rmdirSync(path)
    }
  } else {
    console.log("Directory path not found.")
  }
}


// All of the Node.js APIs are available in the preload process.
// It has the same sandbox as a Chrome extension.
window.addEventListener('DOMContentLoaded', () => {
  
  var btnExcel = document.getElementById("spread");
  var btnOutput = document.getElementById("loc");
  var btnRMexcel = document.getElementById("rmvList");
  var btnRMout = document.getElementById("rmvOut")
  var btnStart = document.getElementById("start");

  var list = document.getElementById("list");
  var output = document.getElementById("output");

  btnRMexcel.onclick = () => {
    list.innerHTML = "";
  }

  btnRMout.onclick = () => {
    output.value = "";
  }

  btnExcel.onclick = ()=>{
    var sheets = dialog.showOpenDialogSync( {
      properties: ['openFile', 'openDirectory', 'multiSelections']
    })

    for(var excel in sheets){
        var option = document.createElement("OPTION");
        var text = document.createTextNode(sheets[excel]);
        option.appendChild(text);
        list.appendChild(option)
    }
  }

  btnOutput.onclick = ()=>{
    var sheets = dialog.showOpenDialogSync( {
     
      properties: ['openFile', 'openDirectory']
    })

    for(var excel in sheets){
      
        output.value = sheets[excel]
    }
  }

  btnStart.onclick =  () => {
    var hold = [];
    var lng = list.options.length;
    var extractToDirectory = output.value;
    
    for(var i = 0; i < lng; i++){
      
      //Extract the name of the spreadsheet...
      var names = list.options[i].text.split("/")
      var position = names[names.length - 1].length;
      var final = names[names.length - 1].substring(0,position - 5);
      fs.mkdirSync(`${extractToDirectory}/${final}`);
      hold.push({"path": extractToDirectory, "folder": final, "excel" : list.options[i].text});
  }


  //Export Both File 
  exportSpreadSheet(hold);


  sleep(500).then(() => {
    deleteUN(hold);
    deleteUN2(hold);
    renamePICS(hold);
    finalCleanUp(hold);
    shell.openItem(extractToDirectory)

  })

  



}

})

function exportSpreadSheet(arr){
    var lng = arr.length;

    for(var i = 0; i < lng; i++){
          //Extract the excel to the desire location...
      fs.createReadStream(arr[i].excel)
      .pipe(unzipper.Extract({ path: `${arr[i].path}/${arr[i].folder}` }))
    }

}


 function deleteUN(arr){
   
    var lng = arr.length;

    for(var i = 0; i < lng; i++){
      
       
        var  items = fs.readdirSync(`${arr[i].path}/${arr[i].folder}`)
        
        items.forEach(item => {
          
        
          if(item.indexOf(".") > 0) {
            fs.unlinkSync(`${arr[i].path}/${arr[i].folder}/${item}`)
          }
          else if(item !=  "xl" ){
            removeDir(`${arr[i].path}/${arr[i].folder}/${item}`)
          }
        })

    }
  }

       
       
        
  function deleteUN2(arr){

    var lng = arr.length;

    for(var i = 0; i < lng; i++){
      var  items = fs.readdirSync(`${arr[i].path}/${arr[i].folder}/xl`)

        items.forEach(item => {
          
        if(item.indexOf(".") > 0) {
          fs.unlinkSync(`${arr[i].path}/${arr[i].folder}/xl/${item}`)
        }
        else if(item !=  "media" ){
          removeDir(`${arr[i].path}/${arr[i].folder}/xl/${item}`)
        }
      })
    }

  }
  
  function renamePICS(arr){
    var lng = arr.length;
    for(var i = 0; i < lng; i++){
      var items = fs.readdirSync(`${arr[i].path}/${arr[i].folder}/xl/media`)
        var index = 0;
        items.forEach(item => {
          console.log(item);
          fs.renameSync(`${arr[i].path}/${arr[i].folder}/xl/media/${item}`,
          `${arr[i].path}/${arr[i].folder}-${(index + 1)}.jpg`)
          index++;
        
      })
    }
    
    
  }
  function finalCleanUp(arr){
    var lng = arr.length;
    for(var i = 0; i < lng; i++){

      fs.rmdirSync(`${arr[i].path}/${arr[i].folder}/xl/media`)
      fs.rmdirSync(`${arr[i].path}/${arr[i].folder}/xl`)
      fs.rmdirSync(`${arr[i].path}/${arr[i].folder}`)

    }

    alert("DONE CHECK LOCATION");
  }

