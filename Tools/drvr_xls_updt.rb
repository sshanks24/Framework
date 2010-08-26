=begin
#------------------------------------------------------------------------------#
This file assumes execution from the \Tools\ folder
Driver Spreadsheet Update -
This script will update the layout  all of the spreadsheets in the ..\driver\
folder. Currently, the updates are based on hardcoding.
#------------------------------------------------------------------------------#
=end
require 'win32ole'

# - create array containing driver path and name of each spreadsheet
# - '$0' is the same as '__FILE__'
def setup
  Dir.chdir(($0).gsub(/Tools.*/,"driver/")) #change path to driver folder
  Dir.entries(Dir.pwd).delete_if{ |e| e=~ /^\..*/|| e=~/^.*\.rb/ ||e=~/backup/i}
end

# - Open spreadsheet and change support layout
def change_xls(file)
  ss = WIN32OLE::new('excel.Application')
  wb = ss.Workbooks.Open(Dir.pwd+'/'+file)
  ws = wb.Worksheets(1)
  ss.visible = true # For debug
  
  row = 9    # from row 9
  # Format if not already formated
  if !((ws.Range("a#{row}").value)=~/[Ss]ummary.*/)
    7.times{|x| ws.Range("a#{row}:b#{row}").Rows.Delete} # delete 7 rows
  
    c = ws.Range("a#{10}:b#{10}")
    c.value = "Item","Value"    # add 'Item' to A10 and 'Value' to B10
    c.Interior.ColorIndex = 20  # change background color
    c.Borders.ColorIndex = 1    # add border
    wb.save
  end
  ss.quit
end

#start script
puts" \n Executing: #{$0}\n\n" # show current filename
setup.each{|x| change_xls(x)} 
puts ' Updates Complete'