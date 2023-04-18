
var tree = {
"id": "root",
"label": "Hauptknoten",
"children": []
};

for (var i = 1; i <= 10; i++) {
var mainNode = {
"id": "mainnode" + i,
"label": "Hauptknoten " + i,
"parent": "root",
"children": []
};

for (var j = 1; j <= 2; j++) {
var subNode = {
"id": "subnode" + i + "_" + j,
"label": "Unterknoten " + i + "." + j,
"parent": "mainnode" + i,
"children": []
};

mainNode.children.push(subNode);

}

tree.children.push(mainNode);
}