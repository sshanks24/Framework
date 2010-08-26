=begin
#------------------------------------------------------------------------------#
Change Driver Spreadsheet IP -
This script expects to be executed for the ..\tools\ folder
It will change the ip address in all of the spreadsheets in ..\driver\ folder.
#------------------------------------------------------------------------------#
=end

require 'win32ole'

# - create instance of excel, open workbook, and write new IP
def chng_xls(xl_name)
  xl = (PATH + xl_name) # ss full name
  puts xl_name # file being updated currently
  ss = WIN32OLE::new('excel.Application')
  ss.visible = true
  wb = ss.Workbooks.Open(xl)
  ws = wb.Worksheets(1)

  # - write new IP to spreadsheet
  ws.Range("b3").value = IP
  wb.save
  ss.quit
end

# - change directory to "..\driver\"
# - backslashes are needed in path to open excel file
Dir.chdir("./driver")
PATH = ((Dir.getwd) + '/').gsub('/','\\')
puts " Driver folder being updated is - #{PATH}"

# - type in new IP address and press enter. or just enter for default address
print "\nEnter the IP Address (press enter for 126.4.202.212): "
IP = gets.chomp!
IP = '126.4.202.212' if IP.empty? # change the default as needed

# - change ip in each file in the array
Dir.glob('*xls').each{|x| chng_xls(x)}
  
puts " \n   Updates Complete"