function traverseFormulaGroups(fGroup) {
    let dataArray = [];
    let linkArray = [];
    let coordToKeys = new Map();
    let keysTorange = new Map();
    fGroup.forEach((formula) => {
        let rangeKey = keysTorange.size;
        formula.loc.key = rangeKey;
        keysTorange.set(rangeKey, formula.loc)
        formula.loc.value.forEach(coord => {
            coordToKeys.set(coord.toString(), [rangeKey]);
        });
    });
    fGroup.forEach((formula) => {
        let operands = formula.operands;
        operands.forEach(operand => {
            let rangeKey = keysTorange.size;
            keysTorange.set(rangeKey, operand)
            operand.key = rangeKey;
            operand.value.forEach(coord => {
                let coordStr = coord.toString();
                if (coordToKeys.has(coordStr)) {
                    coordToKeys.get(coordStr).push(rangeKey);
                } else {
                    coordToKeys.set(coordStr, [rangeKey]);
                }
            });
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
        let overlapMetrics = range.overlapMetrics;
        if (overlapMetrics !== undefined) {
            for (const [overLappingRangeKey, numOverLap] of overlapMetrics.entries()) {
                if (numOverLap === rangeSize) {
                    // overLappingRangeKey and rangeKey are same node
                    let overLappingRange = keysTorange.get(overLappingRangeKey);
                    if (overLappingRange.value.length === rangeSize) {
                        overLappingRange.key = rangeKey;
                        keysTorange.delete(overLappingRangeKey);
                    }
                }
            }
        }
    }
}
function createGraph(fGroup) {
    let dataArray = [];
    let linkArray = [];


    fGroup.forEach((formula) => {
        let cellFormula = formula.cellFormula;
        let operands = formula.operands;

        // Add the node
        dataArray.push({ key: formula.loc.key, name: cellFormula });

        // Add links (parent-child relationships)
        operands.forEach(operand => {
            let opKey = operand.key
            linkArray.push({ from: opKey, to: formula.loc.key });
            if (!dataArray.some(d => d.key === opKey)) {
                dataArray.push({ key: opKey, name: opKey });
            }
        });
    });
    return { nodeDataArray: dataArray, linkDataArray: linkArray };
}
//const json = '[{"cellFormula":"=SUM(RC[-2]:RC[-1])","operands":[[[1,0],[1,1],[2,0],[2,1],[3,0],[3,1],[4,0],[4,1],[5,0],[5,1]]],"loc":[[1,2],[2,2],[3,2],[4,2],[5,2]]},{"cellFormula":"=RC[-2]+RC[-4]","operands":[[[1,2],[2,2],[3,2],[4,2],[5,2]],[[1,0],[2,0],[3,0],[4,0],[5,0]]],"loc":[[1,4],[2,4],[3,4],[4,4],[5,4]]},{"cellFormula":"=R[-6]C-R[-5]C","operands":[[[1,0],[1,1],[1,2]],[[2,0],[2,1],[2,2]]],"loc":[[7,0],[7,1],[7,2]]}]';
const json = '[{"cellFormula":"=SUM(RC[-2]:RC[-1])","operands":[{"value":[[1,0],[1,1],[2,0],[2,1],[3,0],[3,1],[4,0],[4,1],[5,0],[5,1]]}],"loc":{"value":[[1,2],[2,2],[3,2],[4,2],[5,2]]}},{"cellFormula":"=RC[-2]+RC[-4]","operands":[{"value":[[1,2],[2,2],[3,2],[4,2],[5,2]]},{"value":[[1,0],[2,0],[3,0],[4,0],[5,0]]}],"loc":{"value":[[1,4],[2,4],[3,4],[4,4],[5,4]]}},{"cellFormula":"=R[-6]C-R[-5]C","operands":[{"value":[[1,0],[1,1],[1,2]]},{"value":[[2,0],[2,1],[2,2]]}],"loc":{"value":[[7,0],[7,1],[7,2]]}}]';
let fGroup = JSON.parse(json);
traverseFormulaGroups(fGroup);
let ret = createGraph(fGroup);
console.log('hi')