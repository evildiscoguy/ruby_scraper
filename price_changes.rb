# frozen_string_literal: true

require 'open-uri'
require 'nokogiri'
require 'axlsx'
require 'bigdecimal'

# Set the time and date
time = Time.new
date = "#{time.day}/#{time.month}/#{time.year}"

# List of arrays containing all the products by category
garden_room = ['https://www.argos.co.uk/product/9507092',
               'https://www.argos.co.uk/product/9543063',
               'https://www.argos.co.uk/product/9654873',
               'https://www.argos.co.uk/product/9305366']

grey_sofa_living = ['https://www.argos.co.uk/product/9476886',
                    'https://www.argos.co.uk/product/9600243',
                    'https://www.argos.co.uk/product/9444870',
                    'https://www.argos.co.uk/product/9463789',
                    'https://www.argos.co.uk/product/3286869',
                    'https://www.argos.co.uk/product/9637089',
                    'https://www.argos.co.uk/product/9233878',
                    'https://www.argos.co.uk/product/9561887',
                    'https://www.argos.co.uk/product/9581276',
                    'https://www.argos.co.uk/product/9376773']

dining_room = ['https://www.argos.co.uk/product/9397109',
               'https://www.argos.co.uk/product/9376711',
               'https://www.argos.co.uk/product/9627958',
               'https://www.argos.co.uk/product/9619588',
               'https://www.argos.co.uk/product/9174487',
               'https://www.argos.co.uk/product/9548178',
               'https://www.argos.co.uk/product/9534450']

green_sofa_living = ['https://www.argos.co.uk/product/9514731',
                     'https://www.argos.co.uk/product/8957870',
                     'https://www.argos.co.uk/product/1431506',
                     'https://www.argos.co.uk/product/9601510',
                     'https://www.argos.co.uk/product/9547801',
                     'https://www.argos.co.uk/product/9338058',
                     'https://www.argos.co.uk/product/9563603',
                     'https://www.argos.co.uk/product/9417340',
                     'https://www.argos.co.uk/product/9424883',
                     'https://www.argos.co.uk/product/9465000',
                     'https://www.argos.co.uk/product/9598867',
                     'https://www.argos.co.uk/product/9362079']

pidgeon_holes = ['https://www.argos.co.uk/product/9614057',
                 'https://www.argos.co.uk/product/9612183',
                 'https://www.argos.co.uk/product/9558326']

kids_room = ['https://www.argos.co.uk/product/9360882',
             'https://www.argos.co.uk/product/9535363',
             'https://www.argos.co.uk/product/9533231',
             'https://www.argos.co.uk/product/9559507',
             'https://www.argos.co.uk/product/9635012',
             'https://www.argos.co.uk/product/9614655',
             'https://www.argos.co.uk/product/9535363',
             'https://www.argos.co.uk/product/9584682',
             'https://www.argos.co.uk/product/9549421',
             'https://www.argos.co.uk/product/8039365']

bed_room = ['https://www.argos.co.uk/product/5970836',
            'https://www.argos.co.uk/product/9540334',
            'https://www.argos.co.uk/product/4631303',
            'https://www.argos.co.uk/product/9302778',
            'https://www.argos.co.uk/product/5387601',
            'https://www.argos.co.uk/product/9624607',
            'https://www.argos.co.uk/product/9522613',
            'https://www.argos.co.uk/product/4528746',
            'https://www.argos.co.uk/product/8696186']

sony_hardware = ['https://www.argos.co.uk/product/8349000',
                 'https://www.argos.co.uk/product/5718469',
                 'https://www.argos.co.uk/product/6017808',
                 'https://www.argos.co.uk/product/8340322',
                 'https://www.argos.co.uk/product/8349103',
                 'https://www.argos.co.uk/product/8687669']

sony_games = ['https://www.argos.co.uk/product/9482076',
              'https://www.argos.co.uk/product/9377057',
              'https://www.argos.co.uk/product/9509935',
              'https://www.argos.co.uk/product/9482038',
              'https://www.argos.co.uk/product/9481644',
              'https://www.argos.co.uk/product/7442984']

nintendo_hardware = ['https://www.argos.co.uk/product/6851387',
                     'https://www.argos.co.uk/product/6556493',
                     'https://www.argos.co.uk/product/6851150',
                     'https://www.argos.co.uk/product/2077921',
                     'https://www.argos.co.uk/product/2078274']

nintendo_games = ['https://www.argos.co.uk/product/6846440',
                  'https://www.argos.co.uk/product/4659282',
                  'https://www.argos.co.uk/product/6847504',
                  'https://www.argos.co.uk/product/9461963',
                  'https://www.argos.co.uk/product/9461970',
                  'https://www.argos.co.uk/product/1148600',
                  'https://www.argos.co.uk/product/8358864',
                  'https://www.argos.co.uk/product/7918865',
                  'https://www.argos.co.uk/product/9482894']

# Function to take all the arrays and print to sheet
def add_to_sheet(sheet, array, row_data_left, row_data, currency)
  agent = 'Mozilla/5.0 (Windows; U; Win 9x 4.90; SG; rv:1.9.2.4) Gecko/20101104 Netscape/9.1.0285'

  array.each do |product|
    doc = Nokogiri::HTML(URI.open(product, 'User-Agent' => agent))

    title = doc.xpath('//*[@id="content"]/main/div[2]/div[2]/div[1]/section[1]/div[1]/h1/span[1]')
    cat_num = doc.xpath('//*[@id="content"]/main/div[2]/div[2]/div[1]/section[1]/div[1]/h1/span[2]')
    price = doc.xpath('//*[@id="content"]/main/div[2]/div[2]/div[1]/section[2]/section/ul/li/h2')

    sheet.add_row [cat_num.text, title.text, price.text.delete('£').to_f],
                  style: [row_data_left, row_data, currency]
  end

  sheet.add_row
end

# Set up a new workbook
p = Axlsx::Package.new
wb = p.workbook

# Set up the styles for the workbook
s = wb.styles
wb_title = s.add_style sz: 16, b: true
row_header = s.add_style alignment: { horizontal: :center }, b: true, sz: 11
row_data = s.add_style sz: 11
row_data_left = s.add_style sz: 11, alignment: { horizontal: :center }
currency = s.add_style(format_code: '£#,##0.00', sz: 11, alignment: { horizontal: :center })

# 'VGS Wall' sheet in the new worksheet
wb.add_worksheet(name: 'VGS Wall') do |sheet|
  sheet.add_row ["Price Changes for #{date}"], style: wb_title
  sheet.add_row
  sheet.add_row ['Cat Num', 'Product', 'Price (£)'], style: row_header

  sheet.add_row ['Sony', '', ''], style: row_header
  add_to_sheet(sheet, sony_hardware, row_data_left, row_data, currency)
  add_to_sheet(sheet, sony_games, row_data_left, row_data, currency)

  sheet.add_row ['Nintendo', '', ''], style: row_header
  add_to_sheet(sheet, nintendo_hardware, row_data_left, row_data, currency)
  add_to_sheet(sheet, nintendo_games, row_data_left, row_data, currency)

  sheet.column_widths 20, 60, 12
end

# 'Furniture Prices' sheet in the new worksheet
wb.add_worksheet(name: 'Furniture Prices') do |sheet|
  sheet.add_row ["Price Changes for #{date}"], style: wb_title
  sheet.add_row
  sheet.add_row ['Cat Num', 'Product', 'Price (£)'], style: row_header

  sheet.add_row ['Garden Room', '', ''], style: row_header
  add_to_sheet(sheet, garden_room, row_data_left, row_data, currency)

  sheet.add_row ['Grey Sofa Living', '', ''], style: row_header
  add_to_sheet(sheet, grey_sofa_living, row_data_left, row_data, currency)

  sheet.add_row ['Dining Room', '', ''], style: row_header
  add_to_sheet(sheet, dining_room, row_data_left, row_data, currency)

  sheet.add_row ['Green Sofa Living', '', ''], style: row_header
  add_to_sheet(sheet, green_sofa_living, row_data_left, row_data, currency)

  sheet.add_row ['Pidgeon Holes', '', ''], style: row_header
  add_to_sheet(sheet, pidgeon_holes, row_data_left, row_data, currency)

  sheet.add_row ['Kids Room', '', ''], style: row_header
  add_to_sheet(sheet, kids_room, row_data_left, row_data, currency)

  sheet.add_row ['Bed Room', '', ''], style: row_header
  add_to_sheet(sheet, bed_room, row_data_left, row_data, currency)

  sheet.column_widths 20, 60, 12
end

# Save workbook to file name
p.serialize 'price_prices.xlsx'
