let defaultProp = {
  'bold' : false ,
  'italic': false,
  'underline': false,
  'font-family' : 'Arial',
  'font-size' : 16,
  'align' : 'center',
  'content' : ''
}

let cellsAttr = {
  "no_of_rows": 1000,
  "no_of_col" : 30 ,
  "cells_data": {}
}


let sheetContainer = document.getElementById("input-cell-container") ;
let colContainer = document.getElementById("column-name-container") ;
let rowContainer = document.getElementById("row-name-container") ;
let boldTextBtn = document.getElementById("bold-text") ;
let italicTextBtn = document.getElementById("italic-text") ;
let unerlineTextBtn = document.getElementById("underline-text") ;
let formText = document.getElementById("edit-section") ;
let textLeft = document.getElementById("text-left-align") ;
let textCenter = document.getElementById("text-center-align") ;
let textRight = document.getElementById("text-right-align") ;
let sheetTitle = document.getElementById("sheet-title") ;
let sheetNames = document.getElementById("sheets-name") ;

let fontType = document.getElementById("font-type") ;
let fontSizeElm = document.getElementById("font-size") ;
let selectedCell ;
let textPos ;
let row ;
let col ;
let selectedCol ;
let excelSheet ;
let spreadsheetName="Untitled spreadsheet" ; 

let selectedSheet;
let sheetCount = 1 ;

function initializeExcel(){
  sheetContainer.innerHTML ='' ;
  const rows = excelSheet[selectedSheet].no_of_rows ;
  const cols = excelSheet[selectedSheet].no_of_col ;

  let cellData = excelSheet[selectedSheet].cells_data ;

  sheetTitle.innerText = spreadsheetName ;

  for(let i = 1;i<cols;i++){
    let n = i ;
    let colName='' ;
    while(n >0){
      let rem = n%26 ;
      if(rem ==0){
        colName = 'Z' + colName ;
        n = Math.floor(n/26) - 1 ;
      }else{
        colName = String.fromCharCode(rem-1 +65) + colName ;
        n = Math.floor(n/26);
      }
    }
    const newDiv = document.createElement("div") ;
    newDiv.className = "colName";
    newDiv.id = "colId-"+colName;
    newDiv.textContent = colName ;
    colContainer.appendChild(newDiv) ;
  }

  for(let i=1;i<rows;i++){
    const newDiv = document.createElement("div") ;
    newDiv.className = "rowName";
    newDiv.id = "rowId-"+i;
    newDiv.textContent = i ;
    rowContainer.appendChild(newDiv) ;
  }

  for(let j=1;j<cols;j++ ){
    const colDiv = document.createElement("div") ;
    for(let i=1;i<rows;i++){
      colDiv.className = "cell-col" ;
      let currObj ;
      if(cellData[i] && cellData[i][j]){
        currObj = JSON.parse(JSON.stringify(cellData[i][j])) ;
      }else{
        currObj = {...defaultProp} ;
      }
      const cellDiv = document.createElement("div");
      cellDiv.className = "cell" ;
      cellDiv.contentEditable = "true" ;
      cellDiv.id = "row-"+i+"-col-"+j ;
      cellDiv.textContent = currObj.content ;
      cellDiv.style.fontFamily = currObj['font-family'] ;
      cellDiv.style.fontSize = currObj['font-size']+ 'px' ;
      if(currObj.bold)
        cellDiv.classList.add('bold-text') ;
      if(currObj.italic)
        cellDiv.classList.add('italic-text') ;
      if(currObj.underline)
        cellDiv.classList.add('underline-text') ;

      colDiv.appendChild(cellDiv) ;
    }
    sheetContainer.appendChild(colDiv) ;
  }
  generateSheetBottom();
}

function generateSheetBottom(){
  let sheetKeys = Object.keys(excelSheet) ;
  sheetNames.innerHTML = '' ;
  sheetKeys.forEach((key)=>{
    let sheetDiv = document.createElement("div") ;
    sheetDiv.innerText = key ;
    sheetDiv.className = "name-block" ;
    sheetDiv.id = key ;
    if(key == selectedSheet)
      sheetDiv.classList.add('selected-sheet') ;
    sheetNames.appendChild(sheetDiv) ;
  }) ;
}

const autoSave = (e)=>{
  let cellText = e.target.innerText;
  updateProperty('content', cellText) ;
  formText.innerText = cellText;
}

const updateProperty = (property, value)=>{
  let cellData = {...defaultProp};
  cellData[property] = value ;
  if(!excelSheet[selectedSheet].cells_data[row]){
    excelSheet[selectedSheet].cells_data[row] = {} ;
    excelSheet[selectedSheet].cells_data[row][col] =  {...cellData} ;
  }else if(!excelSheet[selectedSheet].cells_data[row][col]){
    excelSheet[selectedSheet].cells_data[row][col] =  {...cellData} ;
  }else{
    excelSheet[selectedSheet].cells_data[row][col][property] = value ;
  }
  if(JSON.stringify(excelSheet[selectedSheet].cells_data[row][col]) === JSON.stringify(defaultProp) ) {
    delete excelSheet[selectedSheet].cells_data[row][col] ;
  }
}

const selectCell = (e) =>{
  selectedCell = e.target;
  formText.textContent = selectedCell.innerText ;

  let cellLoc = selectedCell.id.split("-") ;
  row = cellLoc[1] ;
  col = cellLoc[3] ;

  updateAlignState() ;

  if(selectedCell.classList.contains("bold-text"))
    boldTextBtn.classList.add("selected-optn") ;
  else
    boldTextBtn.classList.remove("selected-optn") ;

  if(selectedCell.classList.contains("italic-text"))
    italicTextBtn.classList.add("selected-optn") ;
  else
    italicTextBtn.classList.remove("selected-optn") ;

  if(selectedCell.classList.contains("underline-text"))
    unerlineTextBtn.classList.add("selected-optn") ;
  else
    unerlineTextBtn.classList.remove("selected-optn") ;
    
  let fontFam =  getComputedStyle(selectedCell).fontFamily ;
  let fontSize=  getComputedStyle(selectedCell).fontSize;
  fontType.value = fontFam;
  fontSizeElm.value = parseInt(fontSize);

}

const unSelectCell = (e)=>{

}

const boldText = () =>{
  selectedCell.classList.toggle("bold-text") ;
  boldTextBtn.classList.toggle("selected-optn") ;
  
  if(selectedCell.classList.contains("bold-text"))
    updateProperty('bold', true) ;
  else
    updateProperty('bold', false) ;

}

const italicText =()=>{
  selectedCell.classList.toggle("italic-text") ;
  italicTextBtn.classList.toggle("selected-optn") ;

  if(selectedCell.classList.contains("italic-text"))
    updateProperty('italic', true) ;
  else
    updateProperty('italic', false) ;
}

const underlineText =()=>{
  selectedCell.classList.toggle("underline-text");
  unerlineTextBtn.classList.toggle("selected-optn") ;

  if(selectedCell.classList.contains("underline-text"))
    updateProperty('underline', true) ;
  else
    updateProperty('underline', false) ;
}

const clickOutside = ()=>{
  
}

const formulaChange = (e)=>{
  selectedCell.innerText = e.target.innerText ;
}

const alignLeft =()=>{
  textPos = 'left' ;
  selectedCell.style.justifyContent = textPos;
  updateProperty('align', textPos) ;
  updateAlignState() ;
}

const alignCenter =()=>{
  textPos = 'center' ;
  selectedCell.style.justifyContent = textPos;
  updateProperty('align', textPos) ;
  updateAlignState() ;
}

const alignRight =()=>{
  textPos = 'right' ;
  selectedCell.style.justifyContent = textPos;
  updateProperty('align', textPos) ;
  updateAlignState() ;
}

const updateAlignState = () =>{
  let textAlign = getComputedStyle(selectedCell).justifyContent ;
  if(textAlign == 'center'){
    textCenter.classList.add("selected-optn");
  }
  else{
    textCenter.classList.remove("selected-optn") ;
  }
  
  if(textAlign == 'left'){
    textLeft.classList.add("selected-optn");
  }
  else{
    textLeft.classList.remove("selected-optn") ;
  }

  if(textAlign == 'right'){
    textRight.classList.add("selected-optn");
  }
  else{
    textRight.classList.remove("selected-optn") ;
  }
}

const saveExcel =(e)=>{
  if (e.ctrlKey && e.key === 's') {
    let newSheetData =  {} ;
    newSheetData[spreadsheetName] = excelSheet ;
    const exlData = JSON.stringify(newSheetData) ;
    localStorage.setItem('excelData', exlData) ;
    sheetTitle.innerText = "Untitled spreadsheet" ;
  }
}

const changeFontType=(e)=>{
  let fontFamily = e.target.value ;
  updateProperty('font-family', fontFamily) ;
  selectedCell.style.fontFamily = fontFamily ;
}

const changeFontSize=(e)=>{
  let fontSize = e.target.value ;
  updateProperty('font-size', fontSize) ;
  selectedCell.style.fontSize = fontSize+'px' ;
}

const changeSheetName =(e)=>{
  spreadsheetName= e.target.innerText || "Untitled spreadsheet" ;
}

const addNewSheet =() =>{
  let cellsData = {...cellsAttr} ;
  sheetCount+=1 ;
  let newsheetName = "Sheet "+ sheetCount ;
  excelSheet[newsheetName] = cellsData ;
  generateSheetBottom();
}

const selectSheet =(e)=>{
  selectedSheet = e.target.id ;
  initializeExcel() ;
}

sheetContainer.addEventListener('input', autoSave) ;

sheetTitle.addEventListener('input', changeSheetName) ;

sheetContainer.addEventListener('click', selectCell) ;

sheetContainer.addEventListener('focusout', unSelectCell) ;

boldTextBtn.addEventListener("click", boldText) ;

italicTextBtn.addEventListener("click", italicText) ;

unerlineTextBtn.addEventListener("click", underlineText) ;

document.addEventListener("click", clickOutside) ;

formText.addEventListener('input', formulaChange) ;

textLeft.addEventListener('click', alignLeft) ;

textRight.addEventListener('click', alignRight) ;

textCenter.addEventListener('click', alignCenter) ;

document.addEventListener('keydown', saveExcel) ;

fontType.addEventListener('change', changeFontType) ;

fontSizeElm.addEventListener('change', changeFontSize) ;

sheetNames.addEventListener('click', selectSheet) ;

document.getElementById("add-sheet").addEventListener('click', addNewSheet) ;

const parsedExcel = localStorage.getItem('excelData') ;

let spreadsheetData ;

if(parsedExcel == null){
  let dummyData = {} ;
  let cellsData = {...cellsAttr} ;
  dummyData[spreadsheetName] = {
    "Sheet 1" : cellsData
  }
  spreadsheetData = JSON.parse(JSON.stringify(dummyData)) ;
}else{
  spreadsheetData = JSON.parse(parsedExcel) ;
}
spreadsheetName = Object.keys(spreadsheetData)[0];
excelSheet = spreadsheetData[spreadsheetName];
selectedSheet = Object.keys(excelSheet)[0] ;
initializeExcel() ;