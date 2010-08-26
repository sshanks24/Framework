require 'win32ole'

def setup(file)
  driver_path = (file).gsub(/Tools.*/,'driver/') #get the driver path
  contr_path = (file).gsub(/Tools.*/,'controller/') #get the controller path
  adriver_name = adriver_name= Array.new
  adriver_name = Dir.entries(driver_path ).delete_if{ |e| e=~ /^\..*/|| e=~/^.*\.xls/|| e=~/backup/i} #read the driver folder and write the rb file name to array

  temp = Dir.entries(contr_path ).delete_if{ |e| e=~ /^\..*/|| e=~/^.*\.rb/} #read the controller  folder and write the ss file name to array
  contr_name = contr_path + temp[0] # controller path and name

  param = Array.new 
  param[0..2] = driver_path, adriver_name,contr_name
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

def change_con(file)
  param = setup(file) #param is array include driver_path,adriver_name,contr_name

  ss =new_xls(param[2])
  wb = ss[1]
  ws = ss[2]

  ws.Range("j2:j101").Columns.Delete
  row =2
  puts  param[1]
  puts " \n** #{param[1].length} script names were added to #{param[2]} **\n\n"
  param[1].each{|x|   #add every driver name to ss
    ws.Range("j#{row}")['Value'] =x
    row =row+1
  }

  wb.save
  ss[0].quit
end


###script running
p "Start Running"
 change_con(__FILE__)
p "End Running"
