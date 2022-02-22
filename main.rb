# frozen_string_literal: true

require 'open-uri'
require 'nokogiri'
require 'axlsx'

# Set date to a dd/mm/yyyy format
time = Time.new
date = "#{time.day}/#{time.month}/#{time.year}"

# Set a User Agent to get pass scraping blocker
agent = 'Mozilla/5.0 (Windows; U; Win 9x 4.90; SG; rv:1.9.2.4) Gecko/20101104 Netscape/9.1.0285'

# Array containing all the product links
products = ['https://www.argos.co.uk/product/8349000',
            'https://www.argos.co.uk/product/5718469',
            'https://www.argos.co.uk/product/6017808',
            'https://www.argos.co.uk/product/8340322',
            'https://www.argos.co.uk/product/8349103',
            'https://www.argos.co.uk/product/8687669',
            'https://www.argos.co.uk/product/9482076',
            'https://www.argos.co.uk/product/9377057',
            'https://www.argos.co.uk/product/9509935',
            'https://www.argos.co.uk/product/9482038',
            'https://www.argos.co.uk/product/9481644',
            'https://www.argos.co.uk/product/7442984',
            'https://www.argos.co.uk/product/6851387',
            'https://www.argos.co.uk/product/6556493',
            'https://www.argos.co.uk/product/6851150',
            'https://www.argos.co.uk/product/2077921',
            'https://www.argos.co.uk/product/2078274',
            'https://www.argos.co.uk/product/6846440',
            'https://www.argos.co.uk/product/4659282',
            'https://www.argos.co.uk/product/6847504',
            'https://www.argos.co.uk/product/9461963',
            'https://www.argos.co.uk/product/9461970',
            'https://www.argos.co.uk/product/1148600',
            'https://www.argos.co.uk/product/8358864',
            'https://www.argos.co.uk/product/7918865',
            'https://www.argos.co.uk/product/9482894']

# Set up new Excel workbook
p = Axlsx::Package.new
wb = p.workbook

# Set up styles for the workbook
s = wb.styles
wb_title = s.add_style sz: 14, b: true
row_header = s.add_style alignment: { horizontal: :center }, b: true, sz: 12
row_data = s.add_style sz: 12
row_data_left = s.add_style sz: 12, alignment: { horizontal: :center }

# Set up workbook and fetch the products
wb.add_worksheet(name: 'Price Changes') do |sheet|
  sheet.add_row ["Price Changes for #{date}"], style: wb_title
  sheet.add_row
  sheet.add_row ['Cat Num', 'Product', 'Price (£)'], style: row_header

  # Loop to grab product detail using xpath and add to workbook
  products.each do |product|
    doc = Nokogiri::HTML(URI.open(product, 'User-Agent' => agent))

    title = doc.xpath('//*[@id="content"]/main/div[2]/div[2]/div[1]/section[1]/div[1]/h1/span[1]')
    cat_num = doc.xpath('//*[@id="content"]/main/div[2]/div[2]/div[1]/section[1]/div[1]/h1/span[2]')
    price = doc.xpath('//*[@id="content"]/main/div[2]/div[2]/div[1]/section[2]/section/ul/li/h2')

    sheet.add_row [cat_num.text, title.text, price.text.delete('£').to_f],
                  style: [row_data_left, row_data, row_data_left]
  end

  sheet.add_row
  sheet.add_row ['❤️ by Lee Jackson']
  sheet.column_widths 20, 60, 12
end

# Write the workbook to file
p.serialize 'price_changes.xlsx'
