import * as ExcelJS from "exceljs";



export default function exportExcelFront(luckysheet, name, excelType , fallback) {
    // 1.创建工作簿，可以为工作簿添加属性
    const workbook = new ExcelJS.Workbook()
    // 2.创建表格，第二个参数可以配置创建什么样的工作表
    luckysheet.forEach(function (table) {
        // debugger
        if (table.data.length === 0) return true
        const worksheet = workbook.addWorksheet(table.name)
        const merge = (table.config && table.config.merge) || {}        //合并单元格
        const borderInfo = (table.config && table.config.borderInfo) || {}      //边框
        const columnWidth = (table.config && table.config.columnlen) || {}    //列宽
        const rowHeight = (table.config && table.config.rowlen) || {}      //行高
        const frozen = table.frozen || {}       //冻结
        const rowhidden = (table.config && table.config.rowhidden) || {}    //行隐藏
        const colhidden = (table.config && table.config.colhidden) || {}    //列隐藏
        const filterSelect = table.filter_select || {}    //筛选
        const images = table.images || {}   //图片
        // console.log(table)
        const hide = table.hide;    //工作表 sheet 1隐藏
        if (hide === 1) {
            // 隐藏工作表
            worksheet.state = 'hidden';
        }
        setStyleAndValue(table.data, worksheet)
        setMerge(merge, worksheet)
        setBorder(borderInfo, worksheet)
        setImages(images, worksheet, workbook)
        setColumnWidth(columnWidth, worksheet)
        //行高设置50导出后在ms-excel中打开显示25，在wps-excel中打开显示50这个bug不会修复
        setRowHeight(rowHeight, worksheet, excelType)
        setFrozen(frozen, worksheet)
        setRowHidden(rowhidden, worksheet)
        setColHidden(colhidden, worksheet)
        setFilter(filterSelect, worksheet)
        return true
    })

    // 4.写入 buffer
    const buffer = workbook.xlsx.writeBuffer().then(data => {
        const blob = new Blob([data], {
            type: 'application/vnd.ms-excel;charset=utf-8'
        })
        console.log("导出成功！")
        fallback(blob, `${name}.xlsx`)
    })
    return buffer
}

/**
 * 列宽
 * @param columnWidth
 * @param worksheet
 */
var setColumnWidth = function (columnWidth, worksheet) {
    for (let key in columnWidth) {
        worksheet.getColumn(parseInt(key) + 1).width = columnWidth[key] / 7.5
    }
}

/**
 * 行高
 * @param rowHeight
 * @param worksheet
 * @param excelType
 */
var setRowHeight = function (rowHeight, worksheet, excelType) {
    //导出的文件用wps打开和用excel打开显示的行高大一倍
    if (excelType == "wps") {
        for (let key in rowHeight) {
            worksheet.getRow(parseInt(key) + 1).height = rowHeight[key] * 0.75
        }
    }
    if (excelType == "office" || excelType == undefined) {
        for (let key in rowHeight) {
            worksheet.getRow(parseInt(key) + 1).height = rowHeight[key] * 1.5
        }
    }
}

/**
 * 合并单元格
 * @param luckyMerge
 * @param worksheet
 */
var setMerge = function (luckyMerge = {}, worksheet) {
    const mergearr = Object.values(luckyMerge)
    mergearr.forEach(function (elem) {
        // elem格式：{r: 0, c: 0, rs: 1, cs: 2}
        // 按开始行，开始列，结束行，结束列合并（相当于 K10:M12）
        worksheet.mergeCells(
            elem.r + 1,
            elem.c + 1,
            elem.r + elem.rs,
            elem.c + elem.cs
        )
    })
}

/**
 * 设置边框
 * @param luckyBorderInfo
 * @param worksheet
 */
var setBorder = function (luckyBorderInfo, worksheet) {
    if (!Array.isArray(luckyBorderInfo)) return

    //合并边框信息
    var mergeCellBorder = function (border1, border2) {
        if (undefined === border1 || Object.keys(border1).length === 0) return border2;
        return Object.assign({}, border1, border2)
    }

    // console.log('luckyBorderInfo', luckyBorderInfo)
    luckyBorderInfo.forEach(function (elem) {
        // 现在只兼容到borderType 为range的情况
        // console.log('ele', elem)
        if (elem.rangeType === 'range') {
            let border = borderConvert(elem.borderType, elem.style, elem.color)
            let rang = elem.range[0]
            let row = rang.row
            let column = rang.column

            let rowBegin = row[0]
            let rowEnd = row[1]
            let colBegin = column[0]
            let colEnd = column[1]
            //处理外边框的情况 没有直接对应的外边框 需要转换成上下左右
            if (border.all) {//全部边框
                let b = border.all
                for (let i = row[0] + 1; i <= row[1] + 1; i++) {
                    for (let y = column[0] + 1; y <= column[1] + 1; y++) {
                        let border = {}
                        border['top'] = b;
                        border['bottom'] = b;
                        border['left'] = b;
                        border['right'] = b;
                        worksheet.getCell(i, y).border = border
                        // console.log(i, y, worksheet.getCell(i, y).border)
                    }
                }
            } else if (border.top) {//上边框
                let b = border.top
                let i = row[0] + 1;
                for (let y = column[0] + 1; y <= column[1] + 1; y++) {
                    let border = {}
                    border['top'] = b;
                    worksheet.getCell(i, y).border = border
                    // console.log(i, y, worksheet.getCell(i, y).border)
                }
            } else if (border.right) {//右边框
                let b = border.right
                for (let i = row[0] + 1; i <= row[1] + 1; i++) {
                    let y = column[1] + 1;
                    let border = {}
                    border['right'] = b;
                    worksheet.getCell(i, y).border = border
                    // console.log(i, y, worksheet.getCell(i, y).border)
                }
            } else if (border.bottom) {//下边框
                let b = border.bottom
                let i = row[1] + 1;
                for (let y = column[0] + 1; y <= column[1] + 1; y++) {
                    let border = {}

                    border['bottom'] = b;
                    worksheet.getCell(i, y).border = border
                    // console.log(i, y, worksheet.getCell(i, y).border)
                }
            } else if (border.left) {//左边框
                let b = border.left
                for (let i = row[0] + 1; i <= row[1] + 1; i++) {
                    let y = column[0] + 1;
                    let border = {}
                    border['left'] = b;
                    worksheet.getCell(i, y).border = border
                    // console.log(i, y, worksheet.getCell(i, y).border)
                }
            } else if (border.outside) {//外边框
                let b = border.outside
                for (let i = row[0] + 1; i <= row[1] + 1; i++) {
                    for (let y = column[0] + 1; y <= column[1] + 1; y++) {
                        let border = {}
                        if (i === rowBegin + 1) {
                            border['top'] = b
                        }
                        if (i === rowEnd + 1) {
                            border['bottom'] = b
                        }
                        if (y === colBegin + 1) {
                            border['left'] = b
                        }
                        if (y === colEnd + 1) {
                            border['right'] = b
                        }
                        let border1 = worksheet.getCell(i, y).border
                        worksheet.getCell(i, y).border = mergeCellBorder(border1, border)
                        // console.log(i, y, worksheet.getCell(i, y).border)
                    }
                }
            } else if (border.inside) {//内边框
                let b = border.inside
                for (let i = row[0] + 1; i <= row[1] + 1; i++) {
                    for (let y = column[0] + 1; y <= column[1] + 1; y++) {
                        let border = {}
                        if (i !== rowBegin + 1) {
                            border['top'] = b
                        }
                        if (i !== rowEnd + 1) {
                            border['bottom'] = b
                        }
                        if (y !== colBegin + 1) {
                            border['left'] = b
                        }
                        if (y !== colEnd + 1) {
                            border['right'] = b
                        }
                        let border1 = worksheet.getCell(i, y).border
                        worksheet.getCell(i, y).border = mergeCellBorder(border1, border)
                        // console.log(i, y, worksheet.getCell(i, y).border)
                    }
                }
            } else if (border.horizontal) {//内侧水平边框
                let b = border.horizontal
                for (let i = row[0] + 1; i <= row[1] + 1; i++) {
                    for (let y = column[0] + 1; y <= column[1] + 1; y++) {
                        let border = {}
                        if (i === rowBegin + 1) {
                            border['bottom'] = b
                        } else if (i === rowEnd + 1) {
                            border['top'] = b
                        } else {
                            border['top'] = b
                            border['bottom'] = b
                        }
                        let border1 = worksheet.getCell(i, y).border
                        worksheet.getCell(i, y).border = mergeCellBorder(border1, border)
                        // console.log(i, y, worksheet.getCell(i, y).border)
                    }
                }
            } else if (border.vertical) {//内侧垂直边框
                let b = border.vertical
                for (let i = row[0] + 1; i <= row[1] + 1; i++) {
                    for (let y = column[0] + 1; y <= column[1] + 1; y++) {
                        let border = {}
                        if (y === colBegin + 1) {
                            border['right'] = b
                        } else if (y === colEnd + 1) {
                            border['left'] = b
                        } else {
                            border['left'] = b
                            border['right'] = b
                        }
                        let border1 = worksheet.getCell(i, y).border
                        worksheet.getCell(i, y).border = mergeCellBorder(border1, border)
                        // console.log(i, y, worksheet.getCell(i, y).border)
                    }
                }
            } else if (border.none) {//当luckysheet边框为border-none的时候表示没有边框 则将对应的单元格border清空
                for (let i = row[0] + 1; i <= row[1] + 1; i++) {
                    for (let y = column[0] + 1; y <= column[1] + 1; y++) {
                        worksheet.getCell(i, y).border = {}
                        // console.log(i, y, worksheet.getCell(i, y).border)
                    }
                }
            }
        }
        if (elem.rangeType === 'cell') {
            // col_index: 2
            // row_index: 1
            // b: {
            //   color: '#d0d4e3'
            //   style: 1
            // }
            const {col_index, row_index} = elem.value
            const borderData = Object.assign({}, elem.value)
            delete borderData.col_index
            delete borderData.row_index
            let border = addborderToCell(borderData, row_index, col_index)
            let border1 = worksheet.getCell(row_index + 1, col_index + 1).border;
            worksheet.getCell(row_index + 1, col_index + 1).border = mergeCellBorder(border1, border)
            // console.log(row_index + 1, col_index + 1, worksheet.getCell(row_index + 1, col_index + 1).border)
        }
    })
}


/**
 * 设置带样式的值
 * @param cellArr
 * @param worksheet
 */
var setStyleAndValue = function (cellArr, worksheet) {
    if (!Array.isArray(cellArr)) return
    cellArr.forEach(function (row, rowid) {
        row.every(function (cell, columnid) {
            if (!cell) return true
            let fill = fillConvert(cell.bg)

            let font = fontConvert(
                cell.ff,
                cell.fc,
                cell.bl,
                cell.it,
                cell.fs,
                cell.cl,
                cell.un
            )
            let alignment = alignmentConvert(cell.vt, cell.ht, cell.tb, cell.tr)
            let value = ''

            if (cell.f) {
                value = {formula: cell.f, result: cell.v}
            } else if (!cell.v && cell.ct && cell.ct.s) {
                // xls转为xlsx之后，内部存在不同的格式，都会进到富文本里，即值不存在与cell.v，而是存在于cell.ct.s之后
                let richText = [];
                let cts = cell.ct.s
                for (let i = 0; i < cts.length; i++) {
                    let rt = {
                        text: cts[i].v,
                        font: fontConvert(cts[i].ff, cts[i].fc, cts[i].bl, cts[i].it, cts[i].fs, cts[i].cl, cts[i].un)
                    }
                    richText.push(rt)
                }
                value = {
                    richText: richText
                };

            } else {
                //设置值为数字格式
                if (cell.v !== undefined && cell.v !== '') {
                    var v = +cell.v;
                    if (isNaN(v)) v = cell.v
                    value = v
                }
            }
            //  style 填入到_value中可以实现填充色
            let letter = createCellPos(columnid)
            let target = worksheet.getCell(letter + (rowid + 1))
            // console.log('1233', letter + (rowid + 1))
            // eslint-disable-next-line no-unused-vars
            for (const key in fill) {
                target.fill = fill
                break
            }
            target.font = font
            target.alignment = alignment
            target.value = value

            try {
                //设置单元格格式
                target.numFmt = cell.ct.fa;
            } catch (e) {
                console.warn(e)
            }

            return true
        })
    })
}

/**
 * 设置图片
 * @param images
 * @param worksheet
 * @param workbook
 */
var setImages = function (images, worksheet, workbook) {
    if (typeof images != "object") return;
    for (let key in images) {
        // console.log(images[key]);
        // "data:image/png;base64,iVBORw0KG..."
        // 通过 base64  将图像添加到工作簿
        const myBase64Image = images[key].src;
        //位置
        const tl = {col: images[key].default.left / 72, row: images[key].default.top / 19}
        // 大小
        const ext = {width: images[key].default.width, height: images[key].default.height}
        const imageId = workbook.addImage({
            base64: myBase64Image,
            //extension: 'png',
        });
        worksheet.addImage(imageId, {
            tl: tl,
            ext: ext
        });
    }
}

/**
 * 冻结行列
 * @param frozen
 * @param worksheet
 */
var setFrozen = function (frozen = {}, worksheet) {
    switch (frozen.type) {
        // 冻结首行
        case 'row': {
            worksheet.views = [
                {state: 'frozen', xSplit: 0, ySplit: 1}
            ];
            break
        }
        // 冻结首列
        case 'column': {
            worksheet.views = [
                {state: 'frozen', xSplit: 1, ySplit: 0}
            ];
            break
        }
        // 冻结行列
        case 'both': {
            worksheet.views = [
                {state: 'frozen', xSplit: 1, ySplit: 1}
            ];
            break
        }
        // 冻结行到选区
        case 'rangeRow': {
            let row = frozen.range.row_focus + 1
            worksheet.views = [
                {state: 'frozen', xSplit: 0, ySplit: row}
            ];
            break
        }
        // 冻结列到选区
        case 'rangeColumn': {
            let column = frozen.range.column_focus + 1
            worksheet.views = [
                {state: 'frozen', xSplit: column, ySplit: 0}
            ];
            break
        }
        // 冻结行列到选区
        case 'rangeBoth': {
            let row = frozen.range.row_focus + 1
            let column = frozen.range.column_focus + 1
            worksheet.views = [
                {state: 'frozen', xSplit: column, ySplit: row}
            ];
        }

    }

}

/**
 * 行隐藏
 * @param rowhidden
 * @param worksheet
 */
var setRowHidden = function (rowhidden = {}, worksheet) {
    for (const key in rowhidden) {
        //如果当前行没有内容则隐藏不生效
        const row = worksheet.getRow(parseInt(key) + 1)
        row.hidden = true;
    }
}

/**
 * 列隐藏
 * @param colhidden
 * @param worksheet
 */
var setColHidden = function (colhidden = {}, worksheet) {
    for (const key in colhidden) {
        const column = worksheet.getColumn(parseInt(key) + 1)
        column.hidden = true;
    }
}

/**
 * 自动筛选器
 * @param filter
 * @param worksheet
 */
var setFilter = function (filter = {}, worksheet) {
    if (Object.keys(filter).length === 0) return
    const from = {
        row: filter.row[0] + 1,
        column: filter.column[0] + 1
    }

    const to = {
        row: filter.row[1] + 1,
        column: filter.column[1] + 1
    }

    worksheet.autoFilter = {
        from: from,
        to: to
    }

}

var fillConvert = function (bg) {
    if (!bg) {
        return {}
    }
    // const bgc = bg.replace('#', '')
    let fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: {argb: bg.startsWith("#") ? bg.replace('#', '') : colorRGBtoHex(bg).replace("#", "")},
    }
    return fill;
}

var fontConvert = function (
    ff = 0,
    fc = '#000000',
    bl = 0,
    it = 0,
    fs = 10,
    cl = 0,
    ul = 0
) {
    // luckysheet：ff(样式), fc(颜色), bl(粗体), it(斜体), fs(大小), cl(删除线), ul(下划线)
    const luckyToExcel = {
        0: '微软雅黑',
        1: '宋体（Song）',
        2: '黑体（ST Heiti）',
        3: '楷体（ST Kaiti）',
        4: '仿宋（ST FangSong）',
        5: '新宋体（ST Song）',
        6: '华文新魏',
        7: '华文行楷',
        8: '华文隶书',
        9: 'Arial',
        10: 'Times New Roman ',
        11: 'Tahoma ',
        12: 'Verdana',
        num2bl: function (num) {
            return num !== 0
        }
    }
    // 出现Bug，导入的时候ff为luckyToExcel的val

    let font = {
        name: typeof ff === 'number' ? luckyToExcel[ff] : ff,
        family: 1,
        size: fs,
        color: {argb: fc.startsWith("#") ? fc.replace('#', '') : colorRGBtoHex(fc).replace("#", "")},
        bold: luckyToExcel.num2bl(bl),
        italic: luckyToExcel.num2bl(it),
        underline: luckyToExcel.num2bl(ul),
        strike: luckyToExcel.num2bl(cl)
    }

    return font
}

var alignmentConvert = function (
    vt = 'default',
    ht = 'default',
    tb = 'default',
    tr = 'default'
) {
    // luckysheet:vt(垂直), ht(水平), tb(换行), tr(旋转)
    const luckyToExcel = {
        vertical: {
            0: 'middle',
            1: 'top',
            2: 'bottom',
            default: 'middle'
        },
        horizontal: {
            0: 'center',
            1: 'left',
            2: 'right',
            default: 'center'
        },
        wrapText: {
            0: false,
            1: false,
            2: true,
            default: false
        },
        textRotation: {
            0: 0,
            1: 45,
            2: -45,
            3: 'vertical',
            4: 90,
            5: -90,
            default: 0
        }
    }

    let alignment = {
        vertical: luckyToExcel.vertical[vt],
        horizontal: luckyToExcel.horizontal[ht],
        wrapText: luckyToExcel.wrapText[tb],
        textRotation: luckyToExcel.textRotation[tr]
    }
    return alignment
}

var borderConvert = function (borderType, style = 1, color = '#000') {
    // 对应luckysheet的config中borderinfo的的参数
    if (!borderType) {
        return {}
    }
    const luckyToExcel = {
        type: {
            'border-all': 'all',
            'border-top': 'top',
            'border-right': 'right',
            'border-bottom': 'bottom',
            'border-left': 'left',
            'border-outside': 'outside',
            'border-inside': 'inside',
            'border-horizontal': 'horizontal',
            'border-vertical': 'vertical',
            'border-none': 'none',
        },
        style: {
            0: 'none',
            1: 'thin',
            2: 'hair',
            3: 'dotted',
            4: 'dashDot', // 'Dashed',
            5: 'dashDot',
            6: 'dashDotDot',
            7: 'double',
            8: 'medium',
            9: 'mediumDashed',
            10: 'mediumDashDot',
            11: 'mediumDashDotDot',
            12: 'slantDashDot',
            13: 'thick'
        }
    }
    let border = {}
    border[luckyToExcel.type[borderType]] = {
        style: luckyToExcel.style[style],
        color: {argb: color.replace('#', '')}
    }
    return border
}

function addborderToCell(borders) {
    let border = {}
    const luckyExcel = {
        type: {
            l: 'left',
            r: 'right',
            b: 'bottom',
            t: 'top'
        },
        style: {
            0: 'none',
            1: 'thin',
            2: 'hair',
            3: 'dotted',
            4: 'dashDot', // 'Dashed',
            5: 'dashDot',
            6: 'dashDotDot',
            7: 'double',
            8: 'medium',
            9: 'mediumDashed',
            10: 'mediumDashDot',
            11: 'mediumDashDotDot',
            12: 'slantDashDot',
            13: 'thick'
        }
    }
    // console.log('borders', borders)
    for (const bor in borders) {
        // console.log(bor)
        if (borders[bor].color.indexOf('rgb') === -1) {
            border[luckyExcel.type[bor]] = {
                style: luckyExcel.style[borders[bor].style],
                color: {argb: borders[bor].color.replace('#', '')}
            }
        } else {
            border[luckyExcel.type[bor]] = {
                style: luckyExcel.style[borders[bor].style],
                color: {argb: borders[bor].color}
            }
        }
    }

    return border
}

function createCellPos(n) {
    let ordA = 'A'.charCodeAt(0)

    let ordZ = 'Z'.charCodeAt(0)
    let len = ordZ - ordA + 1
    let s = ''
    while (n >= 0) {
        s = String.fromCharCode((n % len) + ordA) + s

        n = Math.floor(n / len) - 1
    }
    return s
}

//rgb(255,255,255)转16进制 #ffffff
function colorRGBtoHex(color) {
    color = color.replace("rgb", "").replace("(", "").replace(")", "")
    var rgb = color.split(',');
    var r = parseInt(rgb[0]);
    var g = parseInt(rgb[1]);
    var b = parseInt(rgb[2]);
    return "#" + ((1 << 24) + (r << 16) + (g << 8) + b).toString(16).slice(1);
}