$(document).ready(function () {
  toastr.options = {
    'closeButton': true,
    'debug': false,
    'newestOnTop': false,
    'progressBar': false,
    'positionClass': 'toast-top-right',
    'preventDuplicates': false,
    'showDuration': '1000',
    'hideDuration': '1000',
    'timeOut': '5000',
    'extendedTimeOut': '1000',
    'showEasing': 'swing',
    'hideEasing': 'linear',
    'showMethod': 'fadeIn',
    'hideMethod': 'fadeOut',
    "progressBar": true,
  }
})

const host = window.location.hostname

const FileComp = $("#fileExcel")
const btnProses = $("#btnProses")

FileComp.on("change", function (e) {
  e.preventDefault()

  upload()
})

function uniqueId() {
  let d = new Date(),
    month = '' + (d.getMonth() + 1),
    day = '' + d.getDate(),
    year = d.getFullYear().toString().substr(2, 2)
  h = (d.getHours() < 10 ? '0' : '') + d.getHours()
  i = (d.getMinutes() < 10 ? '0' : '') + d.getMinutes()
  s = (d.getSeconds() < 10 ? '0' : '') + d.getSeconds()

  if (month.length < 2)
    month = '0' + month
  if (day.length < 2)
    day = '0' + day

  return [year, month, day, h, i, s].join('')
}

function upload() {
  const regex = /^([a-zA-Z0-9\s_\\.\-:])+(.xls|.xlsx)$/

  const fileName = FileComp.val().toLowerCase()
  const theFile = FileComp[0].files[0]
  if (regex.test(fileName)) {
    if (typeof (FileReader) != "undefined") {
      let reader = new FileReader()

      //For Browsers other than IE.
      if (reader.readAsBinaryString) {
        reader.onload = function (e) {
          print(e.target.result)
        }
        reader.readAsBinaryString(theFile)
      } else {
        //For IE Browser.
        reader.onload = function (e) {
          let data = ""
          const bytes = new Uint8Array(e.target.result)
          for (let i = 0; i < bytes.byteLength; i++) {
            data += String.fromCharCode(bytes[i])
          }
          print(data)
        }
        reader.readAsArrayBuffer(theFile)
      }

      FileComp.val("")

    } else {
      alert("This browser does not support HTML5.")
    }
  } else {
    alert("Please Upload a valid Excel File !")
  }
}

function print(data) {
  const workbook = XLSX.read(data, {
    type: 'binary'
  })

  const firstSheet = workbook.SheetNames[0]

  const sheetData = XLSX.utils.sheet_to_row_object_array(workbook.Sheets[firstSheet])

  const panjangBaris = sheetData.length
  const ganjil = panjangBaris % 2 > 0 ? true : false

  let isi = ``
  if (ganjil) {
    if (panjangBaris > 1) {
      for (let i = 1; i < sheetData.length; i += 2) {
        const prev = i - 1
        const lot1 = sheetData[prev].LOT
        const lot2 = sheetData[i].LOT

        const name1 = uniqueId() + "1" + prev
        const name2 = uniqueId() + "2" + i

        isi +=
          `
          <div class="container">
            <div id="${name1}" class="diqr"><input type="hidden" value="${lot1}" /></div>
            <div id="${name2}" class="diqr"><input type="hidden" value="${lot2}" /></div>
            <table class="text-tbl">
              <tr>
                <td width="50%">${lot1}</td>
                <td width="50%">${lot2}</td>
              </tr>
            </table>
          </div>
          `
      }

      const lot3 = sheetData[panjangBaris - 1].LOT
      const name3 = uniqueId() + "last"
      isi +=
        `
          <div class="container">
            <div id="${name3}" class="diqr"><input type="hidden" value="${lot3}" /></div>
            <table class="text-tbl">
              <tr>
                <td width="50%">${lot3}</td>
                <td width="50%">&nbsp;</td>
              </tr>
            </table>
          </div>
          `
    } else {
      const lot = sheetData[panjangBaris - 1].LOT
      const name = uniqueId() + "single"
      isi +=
        `
          <div class="container">
            <div id="${name}" class="diqr"><input type="hidden" value="${lot}" /></div>
            <table class="text-tbl">
              <tr>
                <td width="50%">${lot}</td>
                <td width="50%">&nbsp;</td>
              </tr>
            </table>
          </div>
          `
    }
  } else {
    if (panjangBaris > 1) {
      for (let i = 1; i < sheetData.length; i += 2) {
        const prev = i - 1
        const lot1 = sheetData[prev].LOT
        const lot2 = sheetData[i].LOT

        const name1 = uniqueId() + "1" + prev
        const name2 = uniqueId() + "2" + i

        isi +=
          `
          <div class="container">
            <div id="${name1}" class="diqr"><input type="hidden" value="${lot1}" /></div>
            <div id="${name2}" class="diqr"><input type="hidden" value="${lot2}" /></div>
            <table class="text-tbl">
              <tr>
                <td>${lot1}</td>
                <td>${lot2}</td>
              </tr>
            </table>
          </div>
          `
      }
    } else {
      toastr.error("File Excel Kosong !")
    }
  }


  let mywindow = window.open(
    "",
    "PRINT QR CODE - PMS MU BY DEP TIK",
    "resizable=yes,width=720,height=" + screen.height
  )
  mywindow.document.write(
    `<!DOCTYPE html><html lang="en"><head><title>PRINT QR CODE - PMS MU BY DEP TIK</title></head><link rel="stylesheet" href="http://${host}/assets/css/print.css" media="all">`
  )
  mywindow.document.write(
    `<script src="http://${host}/assets/js/jquery-3.6.0.min.js"></`
  )
  mywindow.document.write(
    `script><script src="http://${host}/assets/js/jquery-qrcode-0.18.0.min.js"></`
  )
  mywindow.document.write("script><body>")
  mywindow.document.write(isi)
  mywindow.document.write(
    `<script src="http://${host}/assets/js/qr.js"></`
  )
  mywindow.document.write("script></body></html>")
  mywindow.document.close() // necessary for IE >= 10
  mywindow.focus() // necessary for IE >= 10

  let is_chrome = Boolean(window.chrome)
  if (is_chrome) {
    mywindow.onload = function () {
      setTimeout(function () {
        // wait until all resources loaded
        mywindow.print() // change window to winPrint
        mywindow.close() // change window to winPrint
      }, 200)
    }
  } else {
    mywindow.print()
    mywindow.close()
  }

  return true
}

