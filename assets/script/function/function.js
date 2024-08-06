var setColumn = [
    { width: 25 },
    { width: 25 },
    { width: 25 },
    { width: 25 },
    { width: 25 },
]

var merge = [
    { s: { r: 0, c: 3 }, e: { r: 0, c: 4 } },
    { s: { r: 0, c: 0 }, e: { r: 1, c: 0 } },
    { s: { r: 0, c: 1 }, e: { r: 1, c: 1 } },
    { s: { r: 0, c: 2 }, e: { r: 1, c: 2 } },
];
var actionClick = "";
var arrlistGoc = [];
var arrlistSS = [];
var arrExportExcel = [];
var start = () => {
    $('.wrap-content').css({ height: `${window.innerHeight - 129}px` })
    const dropArea = document.querySelector(".box-add-ss"),
        button = dropArea.querySelector(".button-goc"),
        button2 = dropArea.querySelector(".button-ss"),
        input = dropArea.querySelector(".input-goc");
    const fileGoc = dropArea.querySelector(".file-goc");
    const abc = dropArea.querySelector(".file-ssz");
    let file;
    var filename;
    hideNotification();
    button.onclick = () => {
        input.click();
        showLoading()
        actionClick = 1
    };
    button2.onclick = () => {
        input.click();
        showLoading()
        actionClick = 2
    };

    document.querySelector("input").addEventListener("cancel", (evt) => {
        hideloading()
    });
    input.addEventListener("change", function (e) {
        file = e.target.files[0] || "";
        readFileExcel(file)
        var fileName = e.target.files[0].name;
        if (actionClick == 1) {
            fileGoc.innerText = fileName;
        } else {
            abc.innerText = fileName;
        }
    });
}

// **************************** define Function *********************************************88

let showLoading = () => {
    $("#loading").show();
}
let hideloading = () => {
    $("#loading").hide();
}

let readFileExcel = (datas) => {
    readXlsxFile(datas).then((rows) => {
        // console.log("abx:", indexTitle, indexDes, indexName, indexProject)
        if (actionClick == 1) {
            hadleDataFileGoc(rows)
        }
        if (actionClick == 2) {
            hadleDataFileSS(rows)
        }

    }, () => {
        $('#notification').addClass("bg-danger")
        $('#notification').text("import dữ liệu thất bại vui lòng thử lại sau!");
        $('.notification').show();
        hideNotification();
        hideloading();

    });
}

let hideNotification = () => {
    setTimeout(() => {
        $('.notification').hide();
        $('#notification').removeClass("bg-danger");
        $('#notification').removeClass("bg-success")
    }, 1000);
}

let getIndex = (value, arr) => {
    let res = null
    arr.forEach((items, indexAll) => {
        items.forEach((item, index) => {
            if (item?.toString()?.toUpperCase() == value.toUpperCase()) {
                res = {
                    indexAll: indexAll,
                    index: index
                };
                return res;
            }
        })
    })
    return res;

}


let hadleDataFileGoc = (arrs) => {
    hideloading();
    arrlistGoc = [];
    let indexTitle = null;
    let indexDes = null;
    let indexName = null;
    let indexProject = null;
    indexTitle = getIndex("tên", arrs)?.index;
    indexDes = getIndex("Mã số", arrs)?.index;
    indexName = getIndex("Name", arrs)?.index;
    indexProject = getIndex("dự án", arrs)?.index;

    if (indexTitle != null & indexDes != null & indexName != null & indexProject != null) {
        arrs.forEach((item, index) => {
            if (getIndex("tên", arrs)?.indexAll < index) {
                arrlistGoc.push([item[indexTitle], item[indexDes], item[indexName], item[indexProject]])
            }
        })
    }
    // console.log('final:', arrlistGoc)
}

let hadleDataFileSS = (arrs) => {
    arrlistSS = [];
    let indexTitle = null;
    let indexDes = null;
    let indexName = null;
    let indexProject = null;
    indexTitle = getIndex("tên", arrs)?.index;
    indexDes = getIndex("Mã số", arrs)?.index;
    indexName = getIndex("Name", arrs)?.index;
    indexProject = getIndex("dự án", arrs)?.index;

    if (indexTitle != null & indexDes != null & indexName != null & indexProject != null) {
        arrs.forEach((item, index) => {
            if (getIndex("tên", arrs)?.indexAll < index) {
                arrlistSS.push([item[indexTitle], item[indexDes], item[indexName], item[indexProject]])
            }
        })
    }
    // console.log('finalSS:', arrlistSS)
    if (arrlistGoc.length > 0) {
        compareDataTwoFile()
        // console.log("haha")
    }

}

let compareDataTwoFile = () => {
    // console.log("ok", arrlistGoc, arrlistSS)
    arrlistSS.forEach((item1, index1) => {
        let status = true;
        arrlistGoc.forEach((item2, index2)=> {
            if (item1[2] == item2[2]) {
                if (!arrlistGoc[index2][3].includes(item1[3])) {
                    arrlistGoc[index2][3] = `${ arrlistGoc[index2][3]} + ${item1[3]}`
                }
                status = false;
            }
        })
        if (status) {
            arrlistGoc.push(item1)
        }
    })

    // console.log("hahaOK:",arrlistGoc)
    exportListExcel()
}

let exportListExcel = () => {
    arrlistGoc.unshift(['Tên', "Mã số", 'Name', 'Dự án'])
    // console.log("haha")
    var d = new Date();
    let Today = `${d.getDate()}_${d.getMonth() + 1}_${d.getFullYear()}`;
    var wb = XLSX.utils.book_new();
    var ws = XLSX.utils.aoa_to_sheet(arrlistGoc);
    ws['!cols'] = setColumn;
    // ws["!merges"] = merge;
    XLSX.utils.book_append_sheet(wb, ws, "NEW DATA");
    XLSX.writeFile(wb, `Dữ liệu xuất ngày ${Today}.xlsx`, { numbers: XLSX_ZAHL_PAYLOAD, compression: true });
    hideloading()
}