const axios = require('axios')
const xlsx = require('xlsx')
const cheerio = require('cheerio')

// input from input.xlsx
let wb = xlsx.readFile('input.xlsx')
let ws = wb.Sheets['Sheet1']
let products = xlsx.utils.sheet_to_json(ws)

if (products.length == 0) {
  console.log('Product input empty')
  return
}
const tempProducts = []
//fetch data sequentially
;(async () => {
  await Promise.all(
    products.map((product, index) => {
      return axios
        .get(product.az_link)
        .then((response) => {
          if (response.data) {
            const $ = cheerio.load(response.data)
            const price = $('#priceblock_ourprice').text().replaceAll('$', '')
            const processedProduct = { index, price, link: product.az_link }
            tempProducts.push(processedProduct)
            console.log(`${index} -- added`)
          } else {
            console.log('Invalid response')
          }
        })
        .catch((error) => {
          const errorProduct = {
            index,
            price: '---',
            link: '---',
          }
          tempProducts.push(errorProduct)

          if (error.response && error.response.status == 404) {
            console.log(`${index} -- ${product.az_link} not found ***`)
          } else if (error.response && error.response.status == 401) {
            console.log(
              `${index} -- ${product.az_link} authentication error ***`
            )
          } else if (error.isAxiosError) {
            console.log(
              'Error: Client network socket disconnected before secure TLS connection was established'
            )
          } else {
            console.log(error)
          }
        })
    })
  ).catch((error) => console.log(error))
  // create output
  let wb = xlsx.utils.book_new()
  let ws = xlsx.utils.json_to_sheet(tempProducts)
  let ws_name = 'Sheet1'
  let date = new Date()
  date.setHours(date.getHours() + 6)
  let localDate = date.toISOString().slice(0, 19).replaceAll(':', '-')
  let fileName = 'output-' + localDate + '.xlsx'
  xlsx.utils.book_append_sheet(wb, ws, ws_name)
  xlsx.writeFile(wb, fileName)
})()
