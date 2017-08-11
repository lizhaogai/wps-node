// require('winax');
// var con = new ActiveXObject("KWPS.Application",
//     {
//         activate: false, // Allow activate existance object instance, false by default
//         async: true, // Allow asynchronius calls, true by default (for future usage)
//         type: true
//     });

// console.log(1)
// console.log('----------------------------------');
// let Dispatch = con.Documents;
// console.log(Dispatch);
// console.log('----------------------------------');
// let doc = Dispatch.Open('C:/Users/yifeng/Desktop/wps-node/test.docx',false);
// console.log('----------------------------------');
// console.log(doc);
// doc.ExportAsFixedFormat('C:/Users/yifeng/Desktop/doc-transform-1/tttt111.pdf',17);

// doc.Close();
// //Dispatch.Close();
// con.Quit(1);

// //let doc = Dispatch.Invoke("Open", inputFile, false, true).toDispatch();
// console.log(2)

var convertor = require('../index'); 

var iword = 'C:/Users/yifeng/Desktop/kwps/jianli.docx';
var oword = 'C:/Users/yifeng/Desktop/kwps/jianli.pdf';

var iexcel = 'C:/Users/yifeng/Desktop/kwps/TaokeDetail-2017-03-07.xls';
var oexcel = 'C:/Users/yifeng/Desktop/kwps/TaokeDetail-2017-03-07.pdf';

var ippt = 'C:/Users/yifeng/Desktop/kwps/智能鞋简析v0.1.ppt';
var oppt = 'C:/Users/yifeng/Desktop/kwps/智能鞋简析v0.1.pdf';

convertor.convert2Pdf(iword,oword);
convertor.convert2Pdf(iexcel,oexcel);
convertor.convert2Pdf(ippt,oppt);