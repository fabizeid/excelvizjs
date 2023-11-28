/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

import go from 'gojs'
import { tokenize } from 'excel-formula-tokenizer';
/* global console, document, Excel, Office */
var names = {};
function init() {

  // Since 2.2 you can also author concise templates with method chaining instead of GraphObject.make
  // For details, see https://gojs.net/latest/intro/buildingObjects.html
  const $ = go.GraphObject.make;  // for conciseness in defining templates

  let myDiagram =
    new go.Diagram("myDiagramDiv",
      {
        initialAutoScale: go.Diagram.UniformToFill,
        // define the layout for the diagram
        layout: $(go.TreeLayout, { nodeSpacing: 5, layerSpacing: 30, arrangement: go.TreeLayout.ArrangementFixedRoots })
      });

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
          new go.Binding("text", "key"))
      ),
      $("TreeExpanderButton")
    );

  // Define a trivial link template with no arrowhead.
  myDiagram.linkTemplate =
    $(go.Link,
      { selectable: false },
      $(go.Shape));  // the link shape

  // create the model for the DOM tree
  myDiagram.model =
    new go.TreeModel( {
      isReadOnly: true,  // don't allow the user to delete or copy nodes
      // build up the tree in an Array of node data
      nodeDataArray: traverseDom(document.activeElement)
    });
}

// Walk the DOM, starting at document, and return an Array of node data objects representing the DOM tree
// Typical usage: traverseDom(document.activeElement)
// The second and third arguments are internal, used when recursing through the DOM
function traverseDom(node, parentName, dataArray) {
  if (parentName === undefined) parentName = null;
  if (dataArray === undefined) dataArray = [];
  // skip everything but HTML Elements
  if (!(node instanceof Element)) return;
  // Ignore the navigation menus
  if (node.id === "navSide" || node.id === "navTop") return;
  // add this node to the nodeDataArray
  var name = getName(node);
  var data = { key: name, name: name };
  dataArray.push(data);
  // add a link to its parent
  if (parentName !== null) {
    data.parent = parentName;
  }
  // find all children
  var l = node.childNodes.length;
  for (var i = 0; i < l; i++) {
    traverseDom(node.childNodes[i], name, dataArray);
  }
  return dataArray;
}

function get_formula_groups(formulasR1C1,formulasA) {
  let groups = [];
  for (var row = 0; row < formulasR1C1.length; row++) {
    for (var col = 0; col < formulasR1C1[row].length; col++) {
      let cellFormula = formulasR1C1[row][col];
      if (typeof cellFormula === 'string' && cellFormula.startsWith('=')) {
        // Breadth-First Search (BFS)
        let stack = [[row, col]];
        let current_group = [];
        while (stack.length > 0) {
          let [x, y] = stack.pop();
          let cellFormula = formulasR1C1[x][y];
          let cellFormulaA = formulasA[x][y];
          const tokens = tokenize(cellFormulaA);
          let operands = [];
          tokens.forEach(({ value, type, subtype }) => {
            if (type === 'operand') {
              operands.push(value);
            }
          });

          formulasR1C1[x][y] = null;
          current_group.push({ cellFormula: cellFormula, operands: operands,loc:[x,y] });
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

// Give every node a unique name
function getName(node) {
  var n = node.nodeName;
  if (node.id) n = n + " (" + node.id + ")";
  var namenum = n;  // make sure the name is unique
  var i = 1;
  while (names[namenum] !== undefined) {
    namenum = n + i;
    i++;
  }
  names[namenum] = node;
  return namenum;
}

// When a Node is selected, highlight the corresponding HTML element.
function nodeSelectionChanged(node) {
  if (node.isSelected) {
    names[node.data.name].style.backgroundColor = "lightblue";
  } else {
    names[node.data.name].style.backgroundColor = "";
  }
}

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
});

export async function run() {
  //init();
  try {
    await Excel.run(async (context) => {
      /**
       * Insert your Excel code here
       */


      var sheet = context.workbook.worksheets.getActiveWorksheet();
      var usedRange = sheet.getUsedRange();
      usedRange.load('formulasR1C1');
      usedRange.load('formulas');
      
      await context.sync();
      //console.log(`The range address was ${range.address}.`);
      var formulasR1C1 = usedRange.formulasR1C1;
      var formulasA = usedRange.formulas;
      var outputDiv = document.getElementById("formulas-output");
      outputDiv.innerHTML = ''; // Clear previous output
      let groups = get_formula_groups(formulasR1C1,formulasA);

    });
  } catch (error) {
    console.error(error);
  }
}
