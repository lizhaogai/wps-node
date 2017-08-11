require('winax');

module.exports.convert2Pdf = function (inputFile, outputFile) {
    if (inputFile.endsWith('.doc') || inputFile.endsWith('.docx')) {
        return word2Pdf(inputFile, outputFile);
    } else if (inputFile.endsWith('.xls') || inputFile.endsWith('.xlsx')) {
        return excel2Pdf(inputFile, outputFile);
    } else if (inputFile.endsWith('.ppt') || inputFile.endsWith('.pptx')) {
        return ppt2Pdf(inputFile, outputFile);
    } else {
        return 0;
    }

}

function word2Pdf(inputFile, outputFile) {
    var con = new ActiveXObject("KWPS.Application",
        {
            activate: false,
            async: true,
            type: true
        });
    // con.Visible = false;
    let Dispatch = con.Documents;
    let doc = Dispatch.Open(inputFile, false);
    doc.ExportAsFixedFormat(outputFile, 17);
    doc.Close();
    con.Quit(1);
    return 1;
}


function excel2Pdf(inputFile, outputFile) {
    var con = new ActiveXObject("KET.Application",{
            activate: false,
            async: true,
            type: true
        });
    con.Visible = false;
    let Dispatch = con.Workbooks;
    let excel = Dispatch.Open(inputFile, false);
    excel.ExportAsFixedFormat(0,outputFile, 0);
    excel.Close();
    con.Quit();
    return 1;

}

function ppt2Pdf(inputFile, outputFile) {
    var con = new ActiveXObject("KWPP.Application",
        {
            activate: false,
            async: true,
            type: true
        });
    // con.Visible = false;
    let Dispatch = con.Presentations;
    let ppt = Dispatch.Open(inputFile, false);
    ppt.SaveAs(outputFile, 32);
    ppt.Close();
    con.Quit();
    return 1;
}