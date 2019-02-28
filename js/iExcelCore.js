const jsonSerialize = {
        activeSheetIndex: 0,
        sheetCount: 2,
        sheets: [
        {
                name: 'Details',
                defaults: {
                        rowHeight: 20,
                        colWidth: 64,
                        colHeaderRowHeight: 20,
                        rowHeaderColWidth: 40
                },
                columns: [
                {
                        size: 210
                },
                {
                        size: 114
                }],
                rows: [
                {
                        size: 29
                },
                {
                        size: 23
                },
                {
                        size: 23
                },
                {
                        size: 23
                },
                {
                        size: 23
                }],
                frozenRowCount: 0,
                frozenColCount: 0,
                rowCount: 5,
                columnCount: 2,
                data: {
                        name: 'Details',
                        rowCount: 5,
                        colCount: 2,
                        dataTable: {
                                '0': {
                                        '0': {
                                                value: 'New Houston Store Loan Options',
                                                style: '__builtInStyle001'
                                        },
                                        '1': {
                                                style: '__builtInStyle001'
                                        },
                                        rs: 'e'
                                },
                                '1': {
                                        '0': {
                                                value: 'Amount of Loan',
                                                style: '__builtInStyle002'
                                        },
                                        '1': {
                                                value: '300000',
                                                style: '__builtInStyle003'
                                        },
                                        rs: 'e'
                                },
                                '2': {
                                        '0': {
                                                value: 'Period (years)',
                                                style: '__builtInStyle002'
                                        },
                                        '1': {
                                                value: '3',
                                                style: '__builtInStyle002'
                                        },
                                        rs: 'e'
                                },
                                '3': {
                                        '0': {
                                                value: 'Interest rate (per year)',
                                                style: '__builtInStyle002'
                                        },
                                        '1': {
                                                value: 4,
                                                style: '__builtInStyle004'
                                        },
                                        rs: 'e'
                                },
                                '4': {
                                        '0': {
                                                value: 'Payment (per month)',
                                                style: '__builtInStyle002'
                                        },
                                        '1': {
                                                value: '',
                                                formula: 'PMT(B4/12,B3*12,B2)',
                                                style: '__builtInStyle003'
                                        },
                                        rs: 'e'
                                }
                        },
                        _defaultDataNode: {
                                style: {
                                        foreColor: 'Text 1 0',
                                        hAlign: 3,
                                        vAlign: 2,
                                        font: ' 14.7px/17px Arial',
                                        themeFont: 'Body',
                                        locked: true,
                                        wordWrap: false
                                }
                        }
                },
                spans: [
                {
                        row: 2,
                        rowCount: 1,
                        col: 0,
                        colCount: 2
                },
                {
                        row: 13,
                        rowCount: 1,
                        col: 0,
                        colCount: 4
                }],
                selections: [],
                activeRow: 0,
                activeCol: 0,
                _zoomFactor: 1,
                sheetTabColor: '#FFFFFF',
                names: [],
                _index: 0,
                tabSelected: true,
                isVisible: true,
                viewType: 'NORMAL',
                isEnabled: true
        },
        {
                name: 'Payment Info',
                isVisible: true,
                sheetTabColor: '#FFFFFF',
                tabSelected: false,
                isEnabled: false
        }],
        namedStyles: [
        {
                foreColor: 'Background 1 0',
                hAlign: 1,
                vAlign: 2,
                font: 'normal bold 14px Gill Sans MT',
                themeFont: 'Body',
                locked: true,
                wordWrap: true,
                name: '__builtInStyle001'
        },
        {
                foreColor: 'Background 1 0',
                hAlign: 1,
                vAlign: 2,
                font: 'normal 11px Gill Sans MT',
                themeFont: 'Body',
                formatter: '_("$"* #,##0_);_("$"* \\(#,##0\\);_("$"* "-"??_);_(@_)',
                locked: true,
                wordWrap: true,
                name: '__builtInStyle002'
        },
        {
                foreColor: 'Background 1 0',
                hAlign: 1,
                vAlign: 2,
                font: 'normal 11px Gill Sans MT',
                themeFont: 'Body',
                formatter: '_("$"* #,##0_);_("$"* \\(#,##0\\);_("$"* "-"??_);_(@_)',
                locked: true,
                wordWrap: true,
                name: '__builtInStyle003'
        },
        {
                foreColor: 'Text 1 0',
                hAlign: 1,
                vAlign: 2,
                font: 'normal 11px Gill Sans MT',
                themeFont: 'Body',
                locked: true,
                wordWrap: true,
                name: '__builtInStyle004'
        }],
        theme: 'Office'
}
// Global Initialization
const rows = []
const columns = []
let currentPosition = []
let singleClickCurrentPosition = []
var rowvalue = 212
var columnvalue = 24
let ColMergedCellValue = []
let elePosition = []
RowsandColumCount(jsonSerialize)
/**
 * @description function to calulate mouse position with column and row value
 * @param {Number} value to convert number to character
 * @author Prasanna Brabourame IS5442
 */
function RowsandColumCount(value) {
        try {
                var CoreJson = value
                var rowlength = CoreJson.sheets[0].rows.length
                for (var i = 0; i < rowlength; i++) {
                        rowvalue += CoreJson.sheets[0].rows[i].size
                        rows.push(rowvalue)
                }
                var ColumnLength = CoreJson.sheets[0].columns.length
                for (var i = 0; i < ColumnLength; i++) {
                        columnvalue += CoreJson.sheets[0].columns[i].size
                        columns.push(columnvalue)
                }   
        } catch (error) {
                ErrorLog(error)
        }
}
/**
 * @description function to find top and height calculation for new appeded element
 * @param {event} event to convert number to character
 * @param {Number} type to convert number to character
 * @author Prasanna Brabourame IS5442
 */
function pointerPosition(event, type) {
        try {
                if (document.getElementsByClassName('block').length != 0) {
                        var blockElementLength = document.getElementsByClassName('block').length
                        for (var i = 0; i < blockElementLength; i++) {
                                document.getElementsByClassName('block')[i].remove()
                                formula_status_bar.value = ""
                        }
                }
                currentPosition = []
                ColMergedCellValue = []
                elePosition = []
                var rowPosition, columnPosition
                var cursorX = event.pageX
                var cursorY = event.pageY
                for (var i = 0; i < rows.length; i++) {
                        if (rows[i] >= cursorY) {
                                rowPosition = i
                                break
                        }
                }
                for (var i = 0; i < columns.length; i++) {
                        if (columns[i] >= cursorX) {
                                columnPosition = i
                                break
                        }
                }
                currentPosition.push({
                        row: rowPosition,
                        column: columnPosition
                })
                let rowNumber = rowPosition
                let columnNumber = columnPosition
                getfieldValueonDoubleClick(jsonSerialize, rowNumber, columnNumber, type)
        } catch (error) {
                ErrorLog(error)
        }
        // alert(cursorX+"+"+cursorY);
}

/**
 * @description function to find top and height calculation for new appeded element
 * @param {Number} value to convert number to character
 * @param {Number} row to convert number to character
 * @param {Number} column to convert number to character
 * @param {Number} type to convert number to character
 * @author Prasanna Brabourame IS5442
 */
function getfieldValueonDoubleClick(value, row, column, type) {
        try {
                ColMergedCellValue = []
        let colStartSpan, colEndSpan, eleValue
        let extRow = value.sheets[0].data.dataTable[row]
        let extColumn = extRow[column]
        let extColumnkey = Object.values(extRow)[column]
        let extColunkeyLength = Object.keys(extRow)
        for (var i = column; i < extColunkeyLength.length; i++) {
                if (extColumnkey.hasOwnProperty('value')) {
                        // console.log(extColumnkey.value);
                        colStartSpan = i
                        ColMergedCellValue.push({
                                Start: colStartSpan
                        })
                        eleValue = extColumnkey.value
                        for (var j = colStartSpan; j <= column; j++) {
                                if (Object.values(extRow)[j].hasOwnProperty('value')) {
                                        colEndSpan = j
                                        ColMergedCellValue.push({
                                                End: colEndSpan
                                        })
                                }
                        }
                } else if (extColumnkey.hasOwnProperty('style')) {
                        var colLength = extColunkeyLength.length
                        for (var i = column; i >= 0; i--) {
                                if (Object.values(extRow)[i].hasOwnProperty('value')) {
                                        colStartSpan = i
                                        ColMergedCellValue.push({
                                                Start: colStartSpan
                                        })
                                        eleValue = Object.values(extRow)[i].value
                                        for (var j = i; j < column; j++) {
                                                if (Object.values(extRow)[j].hasOwnProperty('value')) {
                                                        colEndSpan = j + 1
                                                        ColMergedCellValue.push({
                                                                End: colEndSpan
                                                        })
                                                }
                                        }
                                        break
                                }
                        }
                }
                break
        }
        elementAppendCalculation(rows, columns, currentPosition, singleClickCurrentPosition, ColMergedCellValue, eleValue, type)
        } catch (error) {
                ErrorLog(error)
        }
}

/**
 * @description function to find top and height calculation for new appeded element
 * @param {Number} row to convert number to character
 * @param {Number} column to convert number to character
 * @param {Number} currentPosition to convert number to character
 * @param {Number} singleClickCurrentPosition to convert number to character
 * @param {Number} ColMergedCellValue to convert number to character
 * @param {Number} eleValue to convert number to character
 * @param {Number} type to convert number to character
 * @param {Number} CellValue to convert number to character
 * @author Prasanna Brabourame IS5442
 */
function elementAppendCalculation(row, column, currentPosition, singleClickCurrentPosition, ColMergedCellValue, eleValue, type, CellValue) {
        try {
                let heightcalution, top
        if (type == 'DoubleClick') {
                currentPosition = currentPosition
        } else if (type == 'SingleClick') {
                currentPosition = singleClickCurrentPosition
        }
        var initRow = currentPosition[0].row
        if (initRow === 0) {
                top = 212
                heightcalution = Math.abs(top - row[initRow])
        } else {
                top = row[initRow - 1]
                heightcalution = Math.abs(top - row[initRow])
        }
        if (type == 'DoubleClick') {
                colMergedWidthandleft(columns, ColMergedCellValue, top, heightcalution, eleValue, type, null)
        } else if (type == 'SingleClick') {
                colMergedWidthandleft(columns, ColMergedCellValue, top, heightcalution, eleValue, type, CellValue)
        }
        } catch (error) {
                ErrorLog(error)
        }
}
/**
 * @description function to find width and left position of merged cell value
 * @param {Number} columns to convert number to character
 * @param {Number} ColMergedCellValue to convert number to character
 * @param {Number} top to convert number to character
 * @param {Number} heightcalution to convert number to character
 * @param {Number} eleValue to convert number to character
 * @param {Number} type to convert number to character
 * @param {Number} CellValue to convert number to character
 * @author Prasanna Brabourame IS5442
 */
function colMergedWidthandleft(columns, ColMergedCellValue, top, heightcalution, eleValue, type, CellValue) {
        try {
                var left, width
        var startCell = ColMergedCellValue[0].Start
        var EndCell = ColMergedCellValue[1].End
        if (startCell === 0) {
                left = Math.abs(25)
                if (startCell === EndCell) {
                        width = Math.abs(left - 1 - columns[EndCell])
                } else {
                        width = Math.abs(left - 1 - columns[EndCell])
                }
        } else {
                left = Math.abs(columns[startCell - 1])
                if (startCell === EndCell) {
                        width = Math.abs(left - 1 - columns[EndCell])
                } else {
                        width = Math.abs(left - 1 - columns[EndCell])
                }
        }
        elePosition.push({
                top: top,
                left: left,
                width: width,
                heightcalution: heightcalution,
                eleValue: eleValue,
                type: type,
                CellValue: CellValue
        })
        elementAppend('div', elePosition)
        } catch (error) {
                ErrorLog(error)
        }
}
/**
 * @description function to append div with cell value on double click
 * @param {String} tagName which has to append
 * @param {object} Values(5) top,left,width,heightCalculation,eleValue
 * @author Prasanna Brabourame IS5442
 */
function elementAppend(tagName, posObj) {
        try {
                var iDiv = document.createElement('div')
        iDiv.style.position = 'absolute'
        iDiv.style.top = `${posObj[0].top}px`
        iDiv.style.left = `${posObj[0].left}px`
        iDiv.style.width = `${posObj[0].width}px`
        iDiv.style.height = `${posObj[0].heightcalution}px`
        iDiv.style.border = '1px Solid Green'
        if (`${posObj[0].type}` == 'DoubleClick') {
                iDiv.align = 'right'
                iDiv.contentEditable = true
                iDiv.innerHTML = `${posObj[0].eleValue}`
                formula_status_bar.value = `${posObj[0].eleValue}`
                iDiv.id = 'block'
                iDiv.className = 'block'
                iDiv.style.backgroundColor = '#ffffff'
        } else if (`${posObj[0].type}` == 'SingleClick') {
                iDiv.id = `${posObj[0].CellValue}`
                formula_status_bar.value = `${posObj[0].eleValue}`
                iDiv.innerHTML = `<span style="display:none;">${posObj[0].eleValue}</span>`
                iDiv.className = 'block'
        }
        document.getElementsByTagName('body')[0].appendChild(iDiv)
        } catch (error) {
                ErrorLog(error)
        }
}

/**
 * @description function to convert number to variable eg : 1 = A,2 = B,26 = Z,27=AA
 * @param {Number} num to convert number to character
 * @author Prasanna Brabourame IS5442
 */
function toLetters(num) {
        try {
                'use strict'
        var mod = num % 26,
                pow = (num / 26) | 0,
                out = mod ? String.fromCharCode(64 + mod) : (--pow, 'Z')
        return pow ? toLetters(pow) + out : out
        } catch (error) {
                ErrorLog(error)
        }
}
/**
 * @description function to append div with cell value on single click
 * @param {event} e event
 * @param {string} type to define click type single or double
 * @author Prasanna Brabourame IS5442
 */
function cursorPosition(e, type) {
        try {
                if (document.getElementsByClassName('block').length != 0) {
                        var blockElementLength = document.getElementsByClassName('block').length
                        for (var i = 0; i < blockElementLength; i++) {
                                document.getElementsByClassName('block')[i].remove()
                                formula_status_bar.value = ""
                        }
                }
                singleClickCurrentPosition = []
                ColMergedCellValue = []
                elePosition = []
                var rowPosition, columnPosition
                var cursorX = e.pageX
                var cursorY = e.pageY
                for (var i = 0; i < rows.length; i++) {
                        if (rows[i] >= cursorY) {
                                rowPosition = i
                                break
                        }
                }
                for (var i = 0; i < columns.length; i++) {
                        if (columns[i] >= cursorX) {
                                columnPosition = i
                                break
                        }
                }
                singleClickCurrentPosition.push({
                        row: rowPosition,
                        column: columnPosition
                })
                let rowNumber = rowPosition
                let columnNumber = columnPosition
                let CellValue = toLetters(columnPosition + 1) + (rowPosition + 1)
                // alert(toLetters(columnPosition)+rowPosition)
                // alert(
                //   'row' +
                //     singleClickCurrentPosition[0].row +
                //     'colum' +
                //     singleClickCurrentPosition[0].column
                // )
                getfieldValueonSingleClick(jsonSerialize, rowNumber, columnNumber, CellValue, type)
        } catch (error) {
                ErrorLog(error)
        }
}

/**
 * @description function getfieldValueonSingleClick will get the cell value on single mouse click evenet
 * @param {String} value
 * @param {Number} row
 * @param {Number} column
 * @param {Array} CellValue
 * @param {String} type
 * @author Prasanna Brabourame IS5442
 */
function getfieldValueonSingleClick(value, row, column, CellValue, type) {
        try {
                ColMergedCellValue = []
        let colStartSpan, colEndSpan, eleValue
        let extRow = value.sheets[0].data.dataTable[row]
        let extColumn = extRow[column]
        let extColumnkey = Object.values(extRow)[column]
        let extColunkeyLength = Object.keys(extRow)
        for (var i = column; i < extColunkeyLength.length; i++) {
                if (extColumnkey.hasOwnProperty('value')) {
                        // console.log(extColumnkey.value);
                        colStartSpan = i
                        ColMergedCellValue.push({
                                Start: colStartSpan
                        })
                        eleValue = extColumnkey.value
                        for (var j = colStartSpan; j <= column; j++) {
                                if (Object.values(extRow)[j].hasOwnProperty('value')) {
                                        colEndSpan = j
                                        ColMergedCellValue.push({
                                                End: colEndSpan
                                        })
                                }
                        }
                } else if (extColumnkey.hasOwnProperty('style')) {
                        var colLength = extColunkeyLength.length
                        for (var i = column; i >= 0; i--) {
                                if (Object.values(extRow)[i].hasOwnProperty('value')) {
                                        colStartSpan = i
                                        ColMergedCellValue.push({
                                                Start: colStartSpan
                                        })
                                        eleValue = Object.values(extRow)[i].value
                                        for (var j = i; j < column; j++) {
                                                if (Object.values(extRow)[j].hasOwnProperty('value')) {
                                                        colEndSpan = j + 1
                                                        ColMergedCellValue.push({
                                                                End: colEndSpan
                                                        })
                                                }
                                        }
                                        break
                                }
                        }
                }
                break
        }
        elementAppendCalculation(rows, columns, currentPosition, singleClickCurrentPosition, ColMergedCellValue, eleValue, type, CellValue)
        } catch (error) {
                ErrorLog(error)
        }
}

/**
 * @description function to append div with cell value on single click
 * @param {event} e event
 * @author Prasanna Brabourame IS5442
 */
function BuildIt() {
        try {
                fb.innerHTML = '=PMT(' + '<span class="boldblue">' + document.forms[0].input1.value + '</span>' + '%/12,' + '<span class="boldblue">' + document.forms[0].input2.value + '</span>' + '*12,' + '<span class="boldblue">' + document.forms[0].input3.value + '</span>' + ')'
                cf.innerHTML = '=PMT(' + document.forms[0].input1.value + '%/12,' + document.forms[0].input2.value + '*12,' + document.forms[0].input3.value + ')'
                var myIntRate = document.forms[0].input1.value
                var myIntRateText = document.forms[0].input1.value
                myIntRate = parseFloat(myIntRate) / 1200
                var myLoanLength = document.forms[0].input2.value
                var myLoanLengthText = document.forms[0].input2.value
                myLoanLength = parseFloat(myLoanLength).toFixed(0) * 12
                var myLoanAmt = document.forms[0].input3.value
                var myLoanAmtText = document.forms[0].input3.value
                myLoanAmt = parseFloat(myLoanAmt).toFixed(0)
                var paymt = myLoanAmt * Math.pow(1 + myIntRate, myLoanLength) * myIntRate / (Math.pow(1 + myIntRate, myLoanLength) - 1)
                paymt = paymt.toFixed(2)
                document.forms[0].previewBox.value = '$' + paymt
                var formulaForExcel = '=PMT(' + myIntRateText + '%/12,' + myLoanLengthText + '*12,' + myLoanAmtText + ')'
                document.forms[0].formulaText.style.borderStyle = 'dashed'
                document.forms[0].formulaText.style.borderWidth = '2'
                document.forms[0].formulaText.value = formulaForExcel
                document.forms[0].formulaText.select()
                pmtvalue = ('($' + paymt + ')').fontcolor('red')
                setTimeout(function() {
                        cf.innerHTML = pmtvalue
                }, 2000)   
        } catch (error) {
                ErrorLog(error)
        }
};

$('body').on('click', '.pmt_formula_menu', function (e) {
        if($(this).hasClass('active-toggle') == false){
            $(this).removeClass('active-toggle')
        }else{
            $(this).addClass('active-toggle')
        }
    });

/**
 * @description initail autocall function to describe which function to call by single click or double click
 * @param {event} NO_Parameter
 * @author Prasanna Brabourame IS5442
 */
(function() {
        var el
        var count = 0
        var clickCount = 0
        // function ValidatingPMT() {
        //         var elementB5 = document.getElementById('B5')
        //         if (elementB5 != null) {
        //                 elementB5.addEventListener('change click dblclick focusout keydown keypress keyup', function() {
        //                         if (elementB5.innerHTML != '') {
        //                                 var ElementValue = elementB5.innerHTML
        //                                 alert(ElementValue)
        //                         }
        //                 })
        //         }
        // }

        function init(event) {
                try {
                        el = document.getElementById('excel_sheet_1')
                        el.addEventListener('click', function onDblClick(event) {
                                clickCount++
                                if (clickCount === 1) {
                                        singleClickTimer = setTimeout(function() {
                                                clickCount = 0
                                                var type = 'SingleClick'
                                                cursorPosition(event, type)
                                        }, 400)
                                } else if (clickCount === 2) {
                                        count++
                                        clearTimeout(singleClickTimer)
                                        clickCount = 0
                                        var type = 'DoubleClick'
                                        pointerPosition(event, type)
                                }
                        })   
                } catch (error) {
                       ErrorLog(error) 
                } 
        }
        init()
})()

var ErrorLog = function (error) {
        var today = new Date();
        var dd = today.getDate();
        var mm = today.getMonth() + 1; //January is 0!
        var yyyy = today.getFullYear();
        var hh = today.getHours();
        var mmm = today.getMinutes();
        var ss = today.getSeconds();
        if (dd < 10) {
          dd = '0' + dd;
        }
        if (mm < 10) {
          mm = '0' + mm;
        }
        var today = `${dd}/${mm}/${yyyy} - ${hh}:${mmm}:${ss}`;
        console.log('------------------------------------');
        console.log(`--------------${today}--------------`);
        console.log('------------------------------------');
        console.error(`${error}`);
        console.log('------------------------------------');
      }

// UI Based Works for Formula Tab
//  PMT Message
$('body').on('focus', '#pmt_rate_id', function(e) {
        $('.pmt_infor_message').hide()
        $('.pmt_infor_message.rate').show()
})
$('body').on('focus', '#pmt_nper_id', function(e) {
        $('.pmt_infor_message').hide()
        $('.pmt_infor_message.nper').show()
})
$('body').on('focus', '#pmt_pv_id', function(e) {
        $('.pmt_infor_message').hide()
        $('.pmt_infor_message.pv').show()
})
$('body').on('focus', '#pmt_fv__id', function(e) {
        $('.pmt_infor_message').hide()
        $('.pmt_infor_message.fv').show()
})
$('body').on('focus', '#pmt_type_id', function(e) {
        $('.pmt_infor_message').hide()
        $('.pmt_infor_message.type').show()
})

let ErrorFormulae = 'Invalid Input'
//  PMT Value Change
//     $('body').on('keydown', '#pmt_rate_id', function (e) {
//        var get_val = $(this).val();
//        $('.pmt_rate_id_val').addClass('pmt_value_dark');
//        $('.pmt_rate_id_val').text(get_val);
//     });

$('body').on('focusout keypress keydown keyup', '#pmt_rate_id', function(e) {
        var get_val = $(this).val()
        if (get_val.length > 4) {
                if (get_val != `B4/12`) {
                        $('.pmt_rate_id_val').text(ErrorFormulae)
                        // $('.pmt_rate_id_val').css('color','red !important');
                        $('.pmt_rate_id_val').removeClass('pmt_value_dark').addClass('pmt_value_red').removeClass('valueCorrect')
                        $('.pmt_ok_btn').css({
                                'background': 'lightgrey',
                                'cursor': 'not-allowed'
                        });
                } else {
                        $('.pmt_rate_id_val').text(0.003333333)
                        $('.pmt_rate_id_val').removeClass('pmt_value_red').addClass('pmt_value_dark').addClass('valueCorrect')
                        $('.pmt_ok_btn').css({
                                'background': 'lightgrey',
                                'cursor': 'not-allowed'
                        });
                }
        }
        //
        // $('.pmt_rate_id_val').text(get_val);
})

$('body').on('focusout keypress keydown keyup', '#pmt_nper_id', function(e) {
        var get_val = $(this).val()
        if (get_val.length > 4) {
                if (get_val != `B3*12`) {
                        $('.pmt_nper_id_val').text(ErrorFormulae)
                        // $('.pmt_rate_id_val').css('color','red !important');
                        $('.pmt_nper_id_val').removeClass('pmt_value_dark').addClass('pmt_value_red').removeClass('valueCorrect')
                        $('.pmt_ok_btn').css({
                                'background': 'lightgrey',
                                'cursor': 'not-allowed'
                        });
                } else {
                        $('.pmt_nper_id_val').text(36)
                        $('.pmt_nper_id_val').removeClass('pmt_value_red').addClass('pmt_value_dark').addClass('valueCorrect')
                        $('.pmt_ok_btn').css({
                                'background': 'lightgrey',
                                'cursor': 'not-allowed'
                        });
                }
        }
        //
        // $('.pmt_rate_id_val').text(get_val);
})

$('body').on('focusout keypress keydown keyup', '#pmt_pv_id', function(e) {
        var get_val = $(this).val()
        if (get_val.length > 1) {
                if (get_val != `B2`) {
                        $('.pmt_pv_id_val').text(ErrorFormulae)
                        // $('.pmt_rate_id_val').css('color','red !important');
                        $('.pmt_pv_id_val').removeClass('pmt_value_dark').addClass('pmt_value_red').removeClass('valueCorrect')
                        $('.pmt_ok_btn').css({
                                'background': 'lightgrey',
                                'cursor': 'not-allowed'
                        });
                        $('.resultValue').text("number")
                } else {
                        $('.pmt_pv_id_val').text(300000)
                        $('.resultValue').text(-8857.195502)
                        $('.pmt_pv_id_val').removeClass('pmt_value_red').addClass('pmt_value_dark').addClass('valueCorrect')
                        if ($('.valueCorrect').length === 3) {
                                $('.pmt_ok_btn').css({
                                        'background': '#e6e6e6',
                                        'cursor': 'default'
                                });
                                image_generate.style.background = "url('images/stage2.png') no-repeat"
                                var formula_b5 = jsonSerialize.sheets[0].data.dataTable[4][1].formula;
                                jsonSerialize.sheets[0].data.dataTable[4][1].value = formula_b5
                        }
                }
        }
        //
        // $('.pmt_rate_id_val').text(get_val);
})

//     $('body').on('keydown', '#pmt_nper_id', function (e) {
//        var get_val = $(this).val();
//        $('.pmt_nper_id_val').addClass('pmt_value_dark');
//        $('.pmt_nper_id_val').text(get_val);
//     });
//     $('body').on('keydown', '#pmt_pv_id', function (e) {
//        var get_val = $(this).val();
//        $('.pmt_pv_id_val').addClass('pmt_value_dark');
//        $('.pmt_pv_id_val').text(get_val);
//     });

$('body').on('keydown', '#pmt_fv_id', function(e) {
        var get_val = $(this).val()
        $('.pmt_fv_id_val').addClass('pmt_value_dark')
        $('.pmt_fv_id_val').text(get_val)
})
$('body').on('keydown', '#pmt_type_id', function(e) {
        var get_val = $(this).val()
        $('.pmt_type_id_val').addClass('pmt_value_dark')
        $('.pmt_type_id_val').text(get_val)
})
$('.pmt_formula_menu').click(function() {
        $(this).toggleClass('active-toggle')
})
$('.pmt_dialog_box').click(function() {
        $('.pmt_message_box').show()
})
$('body').on('click', '.pmt_ok_btn', function(e) {
        $('.pmt_message_box').hide()
})

$('body').on('click','.pmt_cancel_btn',function(e){
        $('.pmt_message_box').hide()
})
// UI Based Works for Formula Tab Start