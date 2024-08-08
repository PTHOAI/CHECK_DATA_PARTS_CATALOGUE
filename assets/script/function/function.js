var setColumn = [
    { width: 25 },
    { width: 35 },
    { width: 35 },
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
var arrListDesLK = [];
var arrLisTNameVN = [];
var arrListDesBOM = [];
var arrListNameVNBOM = [];
var arrListNameENBOM = [];
var arrListNameMSPHOIBOM = [];
var arrListDesRes = [];
var arrListVNRes = [];
var arrListENRes = [];

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
    console.log("danh mục linh kiện")
    hideloading();
    arrlistGoc = [];
    let indexName = null;
    let indexDes = null;
    indexDes = getIndex("MÃ NVL", arrs)?.index;
    indexName = getIndex("TÊN TIẾNG VIỆT", arrs)?.index;

    if (indexDes != null & indexName != null) {
        arrs.forEach((item, index) => {
            if (item[indexDes] != "MÃ NVL" && item[indexName] != "TÊN TIẾNG VIỆT") {
                arrListDesLK.push(item[indexDes])
                arrLisTNameVN.push(item[indexName])
                // console.log(item)
            }
        })
    }
}

let hadleDataFileSS = (arrs) => {
    console.log("danh mục BOM")
    // console.log("arrListDesLK: ", arrListDesLK)
    // console.log("arrLisTNameVN: ", arrLisTNameVN)
    arrlistSS = [];
    let indexNameEN = null;
    let indexDes = null;
    let indexNamePHOI = null;
    indexNameEN = getIndex("Tên tiếng anh", arrs)?.index;
    indexDes = getIndex("Mã số linh kiện", arrs)?.index;
    indexNamePHOI = getIndex("Mã số phôi", arrs)?.index;

    // console.log("indexDes: ", indexDes)
    // console.log("indexNameEN: ", indexNameEN)
    // console.log("indexNamePHOI: ", indexNamePHOI)
    // console.log("indexNameVN: ", indexDes + 2)
    if (indexDes != null & indexNameEN != null & indexNamePHOI != null) {
        arrs.forEach((item, index) => {
            if (item[indexDes] != null & item[indexDes] != "Mã số linh kiện" & item[indexDes] != "THÔNG TIN VẬT TƯ LINH KIỆN") {
                // console.log("indexNameVN: ", item)
                arrListDesBOM.push(item[indexDes]);
                arrListNameVNBOM.push(item[indexDes + 2]);
                arrListNameENBOM.push(item[indexDes + 3]);
                arrListNameMSPHOIBOM.push(item[indexNamePHOI]);
            }
        })
    }
    // console.log(arrListDesBOM);
    // console.log(arrListNameVNBOM);
    // console.log(arrListNameENBOM);
    // console.log(arrListNameMSPHOIBOM);


    // console.log('finalSS:', arrlistSS)
    if (arrListDesBOM.length > 0) {
        compareDataTwoFile()
    }

}

let compareDataTwoFile = () => {

    arrListDesLK.forEach((item, index) => {
        // index là vị trí mã số danh mục chi tiết
        // getLocationDesBOM(item, arrListDesBOM) vị trí mã số danh mục BOM
        if (getLocation(item, arrListDesBOM) != null) {
            // console.log("item", arrListDesBOM[getLocation(item, arrListDesBOM)])
            // console.log("item", arrListNameVNBOM[getLocation(item, arrListDesBOM)])
            // console.log("item", arrListNameENBOM[getLocation(item, arrListDesBOM)])
            // console.log("item", arrListNameMSPHOIBOM[getLocation(item, arrListDesBOM)], arrListNameMSPHOIBOM)
            arrlistGoc.push([arrListDesBOM[getLocation(item, arrListDesBOM)],arrListNameVNBOM[getLocation(item, arrListDesBOM)],arrListNameENBOM[getLocation(item, arrListDesBOM)],arrListNameMSPHOIBOM[getLocation(item, arrListDesBOM)]])
        } else {
            // mã số không có trong BOM
            // mã số ko có lấy tên tiếng việt để điền tên tiếng anh
            if (getLocation(arrLisTNameVN[index], arrListNameVNBOM) != null) {
                // console.log("item2-MS", item)
                // console.log("item2-VN", arrLisTNameVN[index])
                // console.log("item2-EN", arrListNameENBOM[getLocation(arrLisTNameVN[index], arrListNameVNBOM)])
                // console.log("item2-NCC", "")
                arrlistGoc.push([item,arrLisTNameVN[index],arrListNameENBOM[getLocation(arrLisTNameVN[index], arrListNameVNBOM)],""])
            }else {
                arrlistGoc.push([item,arrLisTNameVN[index],"",""])
            }


        }


    })

    // arrlistGoc.push([1111,2222,3333,4444])
    exportListExcel()
}

let getLocation = (value, arrs) => {
    let res = null
    arrs.forEach((item, index) => {
        if (value == item) {
            res = index
        }
    })
    return res
}


let exportListExcel = () => {
    arrlistGoc.unshift(['Mã số', "Tên tiếng việt", 'Tên tiếng anh', 'NCC'])
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