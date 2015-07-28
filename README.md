to_spreadsheet is a gem that lets you render xls from your existing slim/haml/erb/etc views from Rails (&gt;= 3.0). ![Build Status](https://secure.travis-ci.org/glebm/to_spreadsheet.png?branch=master "Build Status"):http://travis-ci.org/glebm/to_spreadsheet

Installation
------------

Add it to your Gemfile:

    gem 'to_spreadsheet'

Usage
-----

In your controller:

    # my_thingies_controller.rb
    class MyThingiesController < ApplicationController
      respond_to :xls, :html
      def index
        @my_items = MyItem.all
        respond_to do |format|
          format.html 
          format.xlsx { render xlsx: :index, filename: "my_items_doc" }
        end
      end
    end

In your view partial:

    # _my_items.haml
    %table
      %caption My items
      %thead
        %tr
          %td ID
          %td Name
      %tbody
        - my_items.each do |my_item|
          %tr
            %td.number= my_item.id
            %td= my_item.name
      %tfoot
        %tr
          %td(colspan="2") #{my_items.length}

In your index.xls.haml:

    # index.xls.haml
    = render 'my_items', my_items: @my_items

In your index.html.haml:

    # index.html.haml
    = link_to 'Download spreadsheet', my_items_url(format: :xlsx)
    = render 'my_items', my_items: @my_items

### Worksheets

Every table in the view will be converted to a separate sheet.
The sheet title will be assigned to the value of the table’s caption element if it exists.

### Formatting

You can define formats in your view file (local to the view) or in the initializer

    format_xls 'table.my-table' do
      workbook use_autowidth: true
      sheet    orientation: landscape
      format 'th', b: true # bold
      format 'tbody tr', bg_color: lambda { |row| 'ddffdd' if row.index.odd? }
      format 'A3:B10', i: true # italic
      format column: 0, width: 35
      format 'td.custom', lambda { |cell| modify cell somehow.}
      # default value (fallback value when value is blank or 0 for integer / float)
      default 'td.price', 10

For the full list of supported properties head here: http://rubydoc.info/github/randym/axlsx/Axlsx/Styles#add_style-instance_method
In addition, for column formats, Axlsx columnInfo properties are also supported

### Advanced formatting

to_spreadsheet [associates](https://github.com/glebm/to_spreadsheet/blob/master/lib/to_spreadsheet/renderer.rb#L33) HTML nodes with Axlsx objects as follows:

| HTML tag | Axlsx object |
|----------|--------------|
| table    | worksheet    |
| tr       | row          |
| td, th   | cell         |

For example, to directly manipulate a worksheet:

    format_xls do
      format 'table' do |worksheet|
        worksheet.add_chart ...
        # to get the associated Nokogiri node:
        el = context.to_xml_node(worksheet)

### Themes

You can define themes, i.e. blocks of formatting code:

    ToSpreadsheet.theme :zebra do
      format 'tr', bg_color: lambda { |row| 'ddffdd' if row.index.odd? }

And then use them:

    format_xls 'table.zebra', ToSpreadsheet.theme(:zebra)

### Types

The default theme uses class names on td/th to cast values.
Here is the list of class to type mapping:

| CSS class        | Format                   |
|------------------|--------------------------|
| decimal or float | Decimal                  |
| num or int       | Integer                  |
| datetime         | DateTime (Chronic.parse) |
| date             | Date (Date.parse)        |
| time             | Time (Chronic.parse)     |

