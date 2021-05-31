const axios = require('axios')
const xlsx = require('xlsx')
const cheerio = require('cheerio')
const pLimit = require('p-limit')

const limit = pLimit(10)

// input from input.xlsx
let wb = xlsx.readFile('input.xlsx')
let ws = wb.Sheets['Sheet1']
let products = xlsx.utils.sheet_to_json(ws)

if (products.length == 0) {
  console.log('Product input empty')
  return
}
const tempProducts = []
const addProduct = (index, price, link, tempProducts) => {
  const newProduct = { index, price, link }
  tempProducts.push(newProduct)
}
const userAgent =
  'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.77 Safari/537.36'

//fetch data sequentially
;(async () => {
  await Promise.all(
    products.map((product, index) => {
      return limit(() =>
        axios
          .get(product.az_link, { headers: { 'User-Agent': userAgent } })
          .then((response) => {
            if (response.data) {
              const $ = cheerio.load(response.data)
              const price = $('#priceblock_ourprice').text().replaceAll('$', '')
              if (price) {
                addProduct(index, price, product.az_link, tempProducts)
                console.log(`${index} -- added`)
              } else {
                addProduct(index, '---', product.az_link, tempProducts)
                console.log(`${index} -- price unavailable`)
              }
            } else {
              addProduct(index, '---', product.az_link, tempProducts)
              console.log(`${index} -- invalid response`)
            }
          })
          .catch((error) => {
            addProduct(index, '---', product.az_link, tempProducts)
            if (error.response && error.response.status == 404) {
              console.log(`${index} -- not found ***`)
            } else if (error.response && error.response.status == 401) {
              console.log(`${index} -- authentication error ***`)
            } else if (error.isAxiosError) {
              console.log(
                'Error: Client network socket disconnected before secure TLS connection was established'
              )
            } else {
              console.log(error)
            }
          })
      )
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
