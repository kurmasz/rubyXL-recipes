##########################################################################################
#
# Creates a new Excel workbook, populates a few cells with data, then hides column B
#
# (Tested with version 3.4.25.)
#
# (c) 2022 Zachary Kurmas
#
#######################################################################################
require "rubyXL"
require "rubyXL/convenience_methods"

workbook = RubyXL::Workbook.new
sheet = workbook.worksheets.first

sheet.add_cell(0,0, "Sam")
sheet.add_cell(0,1, "George")
sheet.add_cell(0,2, "John")

sheet.cols.get_range(1).hidden = true

workbook.write('hide_b.xlsx')