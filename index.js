const WordExtractor = require("word-extractor");
const extractor = new WordExtractor();
const extracted = extractor.extract("2019-03-08 The Week in Data.docx");

extracted.then(function (doc) {
//   console.log(doc.getBody().split('\n'));
// find all links and get the text above the link

    let startForOutsideWorld  = doc.getBody().split('\n').indexOf('From the Outside World');
    let endForOutsideWorld = doc.getBody().split('\n').indexOf('From the ODI');
    let outsideWorld = doc.getBody().split('\n').slice(startForOutsideWorld, endForOutsideWorld);

    let links = doc.getBody().split('\n').filter(line => line.includes('http'));
    let textAbove = links.map(link => {
        let index = doc.getBody().split('\n').indexOf(link);
        return doc.getBody().split('\n')[index - 1];
    });
    let textBelow = links.map(link => {
        let index = doc.getBody().split('\n').indexOf(link);
        return doc.getBody().split('\n')[index + 1];
    });
    

    // set textAbove as a key, then set the value as an array of objects with the link and textBelow
    let linksObj = {};
    textAbove.forEach((text, i) => {
        linksObj[text] = linksObj[text] || [];
        linksObj[text].push({link: links[i], textBelow: textBelow[i]});
    });
    console.log(linksObj);

    let startForODI  = doc.getBody().split('\n').indexOf('From the ODI');
    let endForODI = doc.getBody().split('\n').indexOf('From the ODI team');
    let odi = doc.getBody().split('\n').slice(startForODI, endForODI);
    
    let odiLnks = doc.getBody().split('\n').filter(line => line.includes('http'));
    let odiTextAbove = odiLnks.map(link => {
        let index = doc.getBody().split('\n').indexOf(link);
        return doc.getBody().split('\n')[index - 1];
    });
    let odiTextBelow = odiLnks.map(link => {
        let index = doc.getBody().split('\n').indexOf(link);
        return doc.getBody().split('\n')[index + 1];
    });

    // set textAbove as a key, then set the value as an array of objects with the link and textBelow
    let odiLinksObj = {};
    odiTextAbove.forEach((text, i) => {
        odiLinksObj[text] = odiLinksObj[text] || [];
        odiLinksObj[text].push({link: odiLnks[i], textBelow: odiTextBelow[i]});
    });

    let start = doc.getBody().split('\n').indexOf('Dear [first name]');
    let end = doc.getBody().split('\n').indexOf('ODI Production Team');
    let body = doc.getBody().split('\n').slice(start, end);
    console.log(...body);
});
