$(".diqr").each(function (i, e) {
  const lot = $(`#${this.id}`).find("input:hidden").val()

  $(`#${this.id}`).qrcode({
    text: lot,
    size: 130,
    minVersion: 1,
    maxVersion: 5,
    radius: 0.5,
    image: "http://simaset.divmu.pindad.co.id/assets/images/airplane.png"
  })

})