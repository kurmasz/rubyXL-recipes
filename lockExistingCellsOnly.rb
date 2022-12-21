##########################################################################################
#
# Creates a new Excel workbook, populates a few cells with data, then
#   1. Protects the populated cells from changes
#   2. Leaves all other cells unprotected.
#
# The challenge is that the unpopulated cells simply don't exist yet, so
# it is not possible to configure them to be unprotected.
# The "trick" is to create a ColumnRange object, set it to cover all columns, 
# then configure it to be unprotected. From what I can tell, the "unprotected"
# style appears to the entire document, except for the cells that are specifically 
# protected.
#
# (Tested with version 3.4.25.)
#
# (c) 2022 Zachary Kurmas
#
#######################################################################################
require "rubyXL"
require "rubyXL/convenience_methods"

# Create a new workbook and grab the default sheet
workbook = RubyXL::Workbook.new
sheet = workbook.worksheets.first

# Crate an xf (formatting record) that locks the cell
locked_xf = workbook.cell_xfs.first.dup
locked_xf.protection = RubyXL::Protection.new(
  locked: true,
  hidden: false,
)
locked_id = workbook.register_new_xf(locked_xf)

# Crate an xf that does not lock the cell
# (I'm not 100% sure this is necessary.)
unlocked_xf = workbook.cell_xfs.first.dup
unlocked_xf.protection = RubyXL::Protection.new(
  locked: false,
  hidden: false,
)
unlocked_id = workbook.register_new_xf(unlocked_xf)

# Create new cells.  Lock each one.
(0..5).each do |row|
  (0..5).each do |col|
    cell = sheet.add_cell(row, col, (row * col).to_s)
    cell.style_index = locked_id
  end
end

# Create a cell range to cover "all" columns. (Upper bound set at 16384)
range = RubyXL::ColumnRange.new
range.min = 1
range.max = 16384
range.width = 10.83203125  # be sure to set this, otherwise columns aren't visible.
range.style_index = unlocked_id # You _may_ be able to simply use the default xf.  I'm not sure.
sheet.cols << range


# Lock the sheet
sheet.sheet_protection = RubyXL::WorksheetProtection.new(
  sheet:          true,
  objects:        true,
  scenarios:      true,
  format_cells:   true,
  format_columns: true,
  insert_columns: true,
  delete_columns: true,
  insert_rows:    true,
  delete_rows:    true
)

workbook.write("lock_test.xlsx")
