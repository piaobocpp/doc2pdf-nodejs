var app = require('http').createServer(handler),
    urllib = require('url'),
    path = require('path'),
    win32ole = require('win32ole'),
    port = 9000;

app.listen(port, function() {
    console.log('Server is listening on port ' + port);
});

function handler(req, res) {
    var params = urllib.parse(req.url, true);
    console.log(params);

    if (params.query && params.query.infile && params.query.outdir) {
        switch(path.extname(params.query.infile)) {
            case '.doc':
            case '.docx':
                res.end(docToPDF(params.query.infile, params.query.outdir));
                break;
            case '.xls':
            case '.xlsx':
                res.end(xlsToPDF(params.query.infile, params.query.outdir));
                break;
            case '.ppt':
            case '.pptx':
                res.end(pptToPDF(params.query.infile, params.query.outdir));
                break;
        }
    }
    else {
        res.end('Hello HuaBei');
    }
}

function docToPDF(inFile, outDir) {
    outPath = '';
    try {
        app = win32ole.client.Dispatch('Word.Application');
        app.Visible = false;
        app.DisplayAlerts = false;

        doc = app.Documents.Open(inFile);
        outPath = path.join(outDir, path.basename(inFile + '.pdf'));
        doc.ExportAsFixedFormat(outPath, 17);
        doc.Saved = true;
        doc.Close(false);

        app.Quit();
    }
    catch(e) {
        console.error(e.message);
    }
    return outPath;
}

function xlsToPDF(inFile, outDir) {
    outPath = '';
    try {
        app = win32ole.client.Dispatch('Excel.Application');
        app.Visible = false;
        app.DisplayAlerts = false;

        xls = app.Workbooks.Open(inFile);
        outPath = path.join(outDir, path.basename(inFile + '.pdf'));
        xls.ExportAsFixedFormat(0, outPath);
        xls.Saved = true;
        xls.Close(false);

        app.Quit();
    }
    catch(e) {
        console.error(e.message);
    }
    return outPath;
}

function pptToPDF(inFile, outDir) {
    outPath = '';
    try {
        app = win32ole.client.Dispatch('PowerPoint.Application');
        app.DisplayAlerts = false;

        ppt = app.Presentations.Open(inFile, true, false, false);
        outPath = path.join(outDir, path.basename(inFile + '.pdf'));
        ppt.SaveAs(outPath, 32);
        ppt.Saved = true;

        app.Quit();
    }
    catch(e) {
        console.error(e.message);
    }
    return outPath;
}
