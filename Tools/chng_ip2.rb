=begin
#------------------------------------------------------------------------------#
Driver Spreadsheet Update -
This script will update the layout  all of the spreadsheets in the ..\driver\
folder. Currently, the updates are based on hardcoding.
#------------------------------------------------------------------------------#
=end


require 'win32ole'


def setup(file)
  driver_path = (file).gsub(/Tools.*/,'driver/') #get the driver path
  afile_name= Array.new
  aexcle_file_name= Dir.entries(driver_path ).delete_if{ |e| e=~ /^\..*/|| e=~/^.*\.rb/  ||e=~/backup/i} #read the driver folder and write the ss file name to array
  param = Array.new # contains no elements. just used as a place holder here
  param[0..1] = driver_path,aexcle_file_name
  return param
end

#  - createand return new instance of excel
def new_xls(s_s) #wb name and sheet number
  ss = WIN32OLE::new('excel.Application')
  wb = ss.Workbooks.Open(s_s)
  ws = wb.Worksheets(1)
  ss.visible = true # For debug
  xls = [ss,wb,ws]

end

def change_xls(excel,driver_name,ip)
  sp = excel[0]+ driver_name #ss whole name
  ss =new_xls(sp)
 
  wb = ss[1]
  ws = ss[2]

  ws.Range("b3").value = ip
  wb.save
  ss[0].quit
  
end



###script running
puts" \n Executing: #{(__FILE__)}\n\n" # show current filename

# Chnage the default ip to your normal test device
puts "Enter the IP Address (press enter for 126.4.202.212):"
ip = gets.chomp!
ip = '126.4.202.212' if ip.empty?


param = setup(__FILE__)
p param
##change every excel information
param[1].each{|x|
  p param,x
  change_xls(param,x,ip)#change one ss

}


puts 'Updates Complete'