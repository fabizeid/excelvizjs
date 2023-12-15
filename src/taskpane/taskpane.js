/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

import go from 'gojs'
import { tokenize } from 'excel-formula-tokenizer';
/* global console, document, Excel, Office */
let names = {};
function init() {

  // Since 2.2 you can also author concise templates with method chaining instead of GraphObject.make
  // For details, see https://gojs.net/latest/intro/buildingObjects.html
  const $ = go.GraphObject.make;  // for conciseness in defining templates

  let myDiagram =
    new go.Diagram("myDiagramDiv",
      {
        initialAutoScale: go.Diagram.UniformToFill,
        // define the layout for the diagram
        layout: $(go.LayeredDigraphLayout, {
          //direction: 90, // layout direction
        layerSpacing: 30, // space between layers
        columnSpacing: 15, // space between columns
          })
      });
// arrangement: go.TreeLayout.ArrangementFixedRoots
  // Define a simple node template consisting of text followed by an expand/collapse button
  myDiagram.nodeTemplate =
    $(go.Node, "Horizontal",
      { selectionChanged: nodeSelectionChanged },  // this event handler is defined below
      $(go.Panel, "Auto",
        $(go.Shape, { fill: "#1F4963", stroke: null }),
        $(go.TextBlock,
          {
            font: "bold 13px Helvetica, bold Arial, sans-serif",
            stroke: "white", margin: 3
          },
          new go.Binding("text", "name"))
      ),
      $("TreeExpanderButton")
    );

  // Define a trivial link template with no arrowhead.
  myDiagram.linkTemplate =
    $(go.Link,
      { selectable: false },
      $(go.Shape));  // the link shape

  return myDiagram;
  }

// When a Node is selected, highlight the corresponding HTML element.
function nodeSelectionChanged(node) {
  if (node.isSelected) {
    highlight(node.data)
  } else {
    clearHighlight(node.data)
  }
}
function updateDiagram(myDiagram, fGroup){
  //downloadObjectAsJson(fGroup, "test")
  //downloadObjectAsJs(fGroup, "test2")
  let linkArray = traverseFormulaGroups(fGroup); //JSON.stringify(fGroup)
  let { nodeDataArray, linkDataArray } = createGraph(fGroup,linkArray)
  myDiagram.model =new go.GraphLinksModel(nodeDataArray, linkDataArray);
}

function downloadObjectAsJson(exportObj, exportName) {
  var dataStr = "data:text/json;charset=utf-8," + encodeURIComponent(JSON.stringify(exportObj));
  var downloadAnchorNode = document.createElement('a');
  downloadAnchorNode.setAttribute("href", dataStr);
  downloadAnchorNode.setAttribute("download", exportName + ".json");
  document.body.appendChild(downloadAnchorNode); // required for Firefox
  downloadAnchorNode.click();
  downloadAnchorNode.remove();
}

function downloadObjectAsJs(exportObj, exportName) {
  // Convert the object to a string and format it as a JS variable
  var jsContent = "const " + exportName + " = " + JSON.stringify(exportObj, null, 4) + ";";

  // Create a Blob object with the JS content and the correct MIME type
  var blob = new Blob([jsContent], { type: 'text/javascript' });
  var url = URL.createObjectURL(blob);

  // Create a temporary link element and trigger the download
  var downloadAnchorNode = document.createElement('a');
  downloadAnchorNode.href = url;
  downloadAnchorNode.download = exportName + ".js";
  document.body.appendChild(downloadAnchorNode); // Required for Firefox
  downloadAnchorNode.click();

  // Clean up by removing the temporary link and revoking the Blob URL
  document.body.removeChild(downloadAnchorNode);
  URL.revokeObjectURL(url);
}
function traverseFormulaGroups(fGroup) {
  let linkArray = [];
  let coordToKeys = new Map();
  let keysTorange = new Map();
  fGroup.forEach((formula) => {
      let rangeKey = keysTorange.size;
      formula.loc.key = rangeKey;
      keysTorange.set(rangeKey, formula.loc)
        if(formula.loc.value.length === 0){
            coordToKeys.set(formula.loc.sheetName, [rangeKey]);
        } else {
      formula.loc.value.forEach(coord => {
          coordToKeys.set(formula.loc.sheetName+coord.toString(), [rangeKey]);
      });
        }
  });
  fGroup.forEach((formula) => {
      let operands = formula.operands;
      operands.forEach(operand => {
          let rangeKey = keysTorange.size;
          keysTorange.set(rangeKey, operand)
          operand.key = rangeKey;
            if(operand.value.length === 0){ 
                let coordStr = operand.sheetName;
                if (coordToKeys.has(coordStr)) {
                    coordToKeys.get(coordStr).push(rangeKey);
                } else {
                    coordToKeys.set(coordStr, [rangeKey]);
                }
            } else {
          operand.value.forEach(coord => {
              let coordStr = operand.sheetName + coord.toString();
              if (coordToKeys.has(coordStr)) {
                  coordToKeys.get(coordStr).push(rangeKey);
              } else {
                  coordToKeys.set(coordStr, [rangeKey]);
              }
          });
            }
      });
  });


  //record overlap keys in each range
  for (const rangeKeys of coordToKeys.values()) {
      if (rangeKeys.length > 1) {
          //there is an overlap
          for (const rangeKey of rangeKeys) {
              let range = keysTorange.get(rangeKey);
              const overlapKeys = range.overlapKeys;
              if (overlapKeys === undefined) {
                  range.overlapKeys = [rangeKeys]
              } else {
                  overlapKeys.push(rangeKeys)
              }
          }
      }
  }

  //process overlap keys to find subsets and matches
  for (const [rangeKey, range] of keysTorange.entries()) {
      let overlapKeys = range.overlapKeys;
      if (overlapKeys !== undefined) {
          let overlapMetrics = new Map();
          for (let coord of overlapKeys) {
              for (let key of coord) {
                  if (rangeKey !== key) {
                      if (overlapMetrics.has(key)) {
                          overlapMetrics.set(key, overlapMetrics.get(key) + 1);
                      } else {
                          overlapMetrics.set(key, 1);
                      }
                  }
              }
          }
          range.overlapMetrics = overlapMetrics;
      }
  }
  for (const [rangeKey, range] of keysTorange.entries()) {
      let rangeSize = range.value.length;
        rangeSize = rangeSize === 0?1:rangeSize;//hack for named ranges
      let overlapMetrics = range.overlapMetrics;
      if (overlapMetrics !== undefined) {
          for (const [overLappingRangeKey, numOverLap] of overlapMetrics.entries()) {
              let overLappingRange = keysTorange.get(overLappingRangeKey);
              if (overLappingRange === undefined){
                  continue; //might have been deleted
              }
              let overLappingRangSize = overLappingRange.value.length;
                overLappingRangSize = overLappingRangSize === 0?1:overLappingRangSize;//hack for named ranges
              if (numOverLap === rangeSize) {
                  // overLappingRangeKey and rangeKey are same node
                  if ( overLappingRangSize === rangeSize) {
                      overLappingRange.key = rangeKey;
                      keysTorange.delete(overLappingRangeKey);
                  } else {
                      //range is a subset of overLappingRange
                      linkArray.push({ from: rangeKey , to:  overLappingRangeKey});
                      //other thatn the above link (and the forlmula link)
                      //don't add any more links to the subsets
                      keysTorange.delete(rangeKey);
                  }
               }
          }
      }
  }
  //remove links from subset nodes (links from their superset should be enough)
  linkArray = linkArray.filter(link => keysTorange.has(link.to));
  for (const [rangeKey, range] of keysTorange.entries()) {
      let rangeSize = range.value.length;
      let overlapMetrics = range.overlapMetrics;
      if (overlapMetrics !== undefined) {
          for (const [overLappingRangeKey, numOverLap] of overlapMetrics.entries()) {
              let overLappingRange = keysTorange.get(overLappingRangeKey);
              if (overLappingRange === undefined){
                  continue; //might have been deleted
              }
              let overLappingRangSize = overLappingRange.value.length;
              if (numOverLap < rangeSize) {
                  if (overLappingRangSize > rangeSize){
                      linkArray.push({ from: rangeKey , to:  overLappingRangeKey });
                  }
              } else {
                  throw new Error('Should never get here');
              }
          }
      }
  }
  return linkArray;
}
function columnNumberToLetter(columnIndex) {
  let columnLetter = '';
  let columnNumber = columnIndex + 1;
  while (columnNumber > 0) {
      let modulo = (columnNumber - 1) % 26;
      columnLetter = String.fromCharCode(65 + modulo) + columnLetter;
      columnNumber = Math.floor((columnNumber - modulo) / 26);
  }
  return columnLetter;
}

function getRangeFromCoord(operand){
  let coordinates = operand.value;
  if (coordinates.length === 0) {
      return null;
  }

  let minX = coordinates[0][0];
  let maxX = coordinates[0][0];
  let minY = coordinates[0][1];
  let maxY = coordinates[0][1];

  coordinates.forEach(coord => {
      minX = Math.min(minX, coord[0]);
      maxX = Math.max(maxX, coord[0]);
      minY = Math.min(minY, coord[1]);
      maxY = Math.max(maxY, coord[1]);
  });
  minX += 1;
  maxX += 1;
  return columnNumberToLetter(minY) + minX  + ':' + columnNumberToLetter(maxY) + maxX;
}
function createGraph(fGroup,linkArray) {
  let dataArray = [];

  fGroup.forEach((formula) => {
      let cellFormula = formula.cellFormula;
      let fName;
      if(formula.loc.value.length === 0){
        fName = 'workbook!' + formula.loc.sheetName + "\n" + cellFormula;
      } else {
        fName = getRangeFromCoord(formula.loc) + "\n" + cellFormula
      }
      // Add the node
      dataArray.push({ key: formula.loc.key, name: fName,range:formula.loc });
    });
    fGroup.forEach((formula) => {
      let operands = formula.operands;
      const uniqueOperands = new Set();
      // Add links (parent-child relationships)
      operands.forEach(operand => {
          let opKey = operand.key
          let name;
          if(operand.value.length === 0){
            name = 'workbook!' + operand.sheetName;
          } else {
            name = getRangeFromCoord(operand);
            if(formula.loc.sheetName !== operand.sheetName){
              name = operand.sheetName + "!" + name;
            }
          }
          if(!uniqueOperands.has(name)){
            uniqueOperands.add(name);
            linkArray.push({ to: opKey, from: formula.loc.key });
            if (!dataArray.some(d => d.key === opKey)) {
                dataArray.push({ key: opKey, name:name ,range:operand });
            }
        }
      });
  });
  return { nodeDataArray: dataArray, linkDataArray: linkArray };
}
function getRangeNamesToRef(rangeNames, rangeNamesw){
 let rangeNamesToRef = new Map()
  rangeNames.forEach(({ name, formula}) => {
    rangeNamesToRef.set(name,formula.substring(1))
  });
  rangeNamesw.forEach(({ name, formula}) => {
    if(!rangeNamesToRef.has(name))
    rangeNamesToRef.set(name,formula.substring(1))
  });
  return rangeNamesToRef;
}
function get_named_groups(activeSheetName,rangeNames, rangeNamesw){
  let groups = [];
  rangeNamesw.forEach(({ name, formula}) => {
    let current_group = {cellFormula: formula, operands:[],loc:{sheetName: name,value:[]}};
    const tokens = tokenize(formula);
    let index = 0;
    tokens.forEach(({ value, type, subtype}) => {
      if (type === 'operand' && subtype === 'range') {
        // Initialize operands[index] with an empty array if it doesn't exist
        current_group.operands[index] ||= {value:[]};
        let [sheetName ,coordValue] = parseReference(activeSheetName,value,[0,0],false);
        current_group.operands[index].value.push(...coordValue);
        current_group.operands[index].sheetName = sheetName;
        index++;
      }
    });
    groups.push(current_group);
  });
  return groups;
}
function get_formula_groups(activeSheetName,rangeNamesToRef,startCoord,formulasR1C1,formulasA) {
  let groups = [];
  for (let row = 0; row < formulasR1C1.length; row++) {
    for (let col = 0; col < formulasR1C1[row].length; col++) {
      let cellFormula = formulasR1C1[row][col];
      if (typeof cellFormula === 'string' && cellFormula.startsWith('=')) {
        // Breadth-First Search (BFS)
        let stack = [[row, col]];
        let current_group = {cellFormula: cellFormula, operands:[],loc:{sheetName: activeSheetName,value:[]}};
        while (stack.length > 0) {
          let [x, y] = stack.pop();
          let cellFormula = formulasR1C1[x][y];
          let cellFormulaA = formulasA[x][y];
          const tokens = tokenize(cellFormula);
          let index = 0;
          tokens.forEach(({ value, type, subtype}) => {
            if (type === 'operand' && subtype === 'range') {
              // Initialize operands[index] with an empty array if it doesn't exist
              current_group.operands[index] ||= {value:[]};
              let [sheetName ,coordValue] = parseReference(activeSheetName,value,[x+startCoord[0],y+startCoord[1]],true);
              current_group.operands[index].value.push(...coordValue);
              current_group.operands[index].sheetName = sheetName;
              index++;
            }
          });

          formulasR1C1[x][y] = null;
          current_group.loc.value.push([x+startCoord[0],y+startCoord[1]]);
          const directions = [[1, 0], [-1, 0], [0, 1], [0, -1]]; // Directions: down, up, right, left
          for (let [dx, dy] of directions) {
            let new_x = x + dx;
            let new_y = y + dy;
            if (0 <= new_x && new_x < formulasR1C1.length && 0 <= new_y && new_y < formulasR1C1[0].length) {
              //console.log(`value:   ${cellFormula}:${formulasR1C1[new_x][new_y]}`);
              if (cellFormula === formulasR1C1[new_x][new_y]) {
                //console.log('same');
                stack.push([new_x, new_y]);
              }
            }
          }
        }
        groups.push(current_group);
      }
    }
  }
  return groups;
}

function parseReference(activeSheetName,ref, baseRC,R1C1) {
  let baseRow = baseRC[0];
  let baseCol = baseRC[1];
  let sheetName = activeSheetName;
  let parseSingleReference;
  if(R1C1){
    parseSingleReference = parseSingleR1C1Reference;
  } else {
    parseSingleReference = parseSingleA1Reference;
  }
  if (ref.includes('!')) {
    let parts = ref.split('!');
    sheetName = parts[0].replace(/^'|'$/g, ''); //remove extra quotes if they exist
    ref = parts[1];
  }
  if (ref.includes(':')) {
      let parts = ref.split(':');
      let start = parseSingleReference(parts[0], baseRow, baseCol);
      let end = parseSingleReference(parts[1], baseRow, baseCol);

      let allCells = [];
      for (let row = start.row; row <= end.row; row++) {
          for (let col = start.column; col <= end.column; col++) {
              allCells.push([row, col]);
          }
      }
      return [sheetName, allCells];
  } else {
      let singleRef = parseSingleReference(ref, baseRow, baseCol);
      if (singleRef == null) {
        //named range
        return [ref.replace(/^@/, ''),[]];
      }
      return [sheetName, [[singleRef.row, singleRef.column]]];
  }
}
function parseSingleA1Reference(ref, baseRow, baseCol) {

  let matchCol = ref.match(/[A-Z]+/)
  let matchRow = ref.match(/\d+/)
  if (matchCol == null || matchRow == null) {
    return null;
  }
  const column = matchCol[0];
  const row = parseInt(matchRow[0], 10);

  let colNumber = 0;
  for (let i = 0; i < column.length; i++) {
      colNumber = colNumber * 26 + (column.charCodeAt(i) - 64);
  }

  return { row: row - 1, column: colNumber-1};
}
function parseSingleR1C1Reference(ref, baseRow, baseCol) {
  let match = ref.match(/R(\[?-?\d*\]?)(?:C(\[?-?\d*\]?))?/);
  if (match == null) {
    return null;
  }
  let rowOffset = match[1];
  let colOffset = match[2];

  rowOffset = rowOffset.includes('[') ? parseInt(rowOffset.replace(/\[|\]/g, '')) : parseInt(rowOffset) || 0;
  colOffset = colOffset.includes('[') ? parseInt(colOffset.replace(/\[|\]/g, '')) : parseInt(colOffset) || 0;

  return { row: rowOffset+baseRow, column: colOffset+baseCol };
}

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    let myDiagram = init();
    document.getElementById("run").onclick = function() {
      run(myDiagram);
    };
  }
});

export async function run(myDiagram) {
  //myDiagram = init();
  try {
    await Excel.run(async (context) => {
      /**
       * Insert your Excel code here
       */


      let sheet = context.workbook.worksheets.getActiveWorksheet();
      let rangeNamesw = context.workbook.names
      let rangeNames = sheet.names;
      rangeNames.load("items/name,items/formula");
      rangeNamesw.load("items/name,items/formula");

      sheet.load('name');
      let usedRange = sheet.getUsedRange();
      usedRange.load('formulasR1C1');
      usedRange.load('formulas');
      usedRange.load('rowIndex');
      usedRange.load('columnIndex');
      //let namedRange = sheet.names.getItem("Mode");
      await context.sync();
      //console.log(`The range address was ${range.address}.`);
      let formulasR1C1 = usedRange.formulasR1C1;
      let formulasA = usedRange.formulas;
      let rangeNamesToRef = getRangeNamesToRef(rangeNames.items, rangeNamesw.items)
      let groupsN = get_named_groups(sheet.name,rangeNames.items, rangeNamesw.items);
      let groups = get_formula_groups(sheet.name, rangeNamesToRef,
      [usedRange.rowIndex, usedRange.columnIndex],
        formulasR1C1,
        formulasA);
      //groups = calculateRC(groups);
      updateDiagram(myDiagram,  groupsN.concat(groups))
    });
  } catch (error) {
    console.error(error);
  }
}

export async function highlight(nodeData) {
  try {
    await Excel.run(async (context) => {
      //console.log(nodeData.key);
      let coordinates = nodeData.range.value;
      let sheetName = nodeData.range.sheetName;
      let sheet;
      if (coordinates.length === 0) {
        nodeData.originalColors = [];
        return;
      }
      if (sheetName === "") {
        sheet = context.workbook.worksheets.getActiveWorksheet();
      } else {
        sheet = context.workbook.worksheets.getItem(sheetName);
        sheet.activate();
      }
      let highlightCells = [];
      coordinates.forEach(coord => {
        let cell = sheet.getCell(coord[0], coord[1]);
        cell.load('format/fill/color');
        highlightCells.push(cell);
      });

      await context.sync();
      let originalColors = [];
      coordinates.forEach((coord , index) => {
        let cell = highlightCells[index];
        originalColors.push({coord: coord, color: cell.format.fill.color});
        cell.format.fill.color = 'yellow';
      });
      nodeData.originalColors = originalColors;

      await context.sync();
    });
  } catch (error) {
    console.error(error);
  }
}

export async function clearHighlight(nodeData) {
  try {
    await Excel.run(async (context) => {
      let sheetName = nodeData.range.sheetName;
      let sheet;
      if(nodeData.originalColors.length === 0){
        return;
      }
      if (sheetName === "") {
        sheet = context.workbook.worksheets.getActiveWorksheet();
      } else {
        sheet = context.workbook.worksheets.getItem(sheetName);
      }
      nodeData.originalColors.forEach(item => {
        let cell = sheet.getCell(item.coord[0], item.coord[1]);
        if(item.color === "#FFFFFF"){
          cell.format.fill.clear();
        } else {
          cell.format.fill.color = item.color;
        }
      });
      await context.sync();
    });
  } catch (error) {
    console.error(error);
  }
}
