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
  afile_name = Array.new
  #read the driver folder and write the ss file name to array
  aexcle_file_name = Dir.entries(driver_path ).delete_if{ |e| e=~ /^\..*/|| e=~/^.*\.rb/  ||e=~/backup/i}
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

  
def change_ver(file,type=nil)

  #change file name with a new version, according to the file type or change all the file in the folder
  #type is postfix of a file ,such as rb , xls, doc
  driver_path = (file).gsub(/Tools.*/,'driver/').gsub('/','\\') #get the driver path
  p driver_path
 Dir.chdir(driver_path)

  puts "What files would you like to change? Enter '1' for ruby, '2' for excel or press enter for both file types"
  f_type = gets.chomp!
  type  = "rb"  if f_type.to_i == 1
  type  = "xls" if f_type.to_i == 2
  type  = nil    if f_type.empty?
  p type
  if type.nil?
    afile_name= Dir.entries(driver_path ).delete_if{ |e|e=~ /^\..*/||e=~/backup/i}
  else
    afile_name= Dir.entries(driver_path ).delete_if{ |e| e=~ /^\..*/ || !(e=~/^.*\.#{type.to_s}/) || e=~/backup/i} 
  end
p afile_name
  #  change the file version
  afile_name.each{|x|
    p x
    if temp = /([0-9])./.match(x)
p "sssss"
      vn = (temp[0].to_i) +1 #vn is version number
      new_filename =  (x).gsub(/[0-9]+/,vn.to_s)
        p "222222222222"
      p x
      p new_filename
      p "222222222222"
      File.rename(x,new_filename)
    end
  }
end

###script running
puts" \n Executing: #{(__FILE__)}\n\n" # show current filename


puts 'Changing Version Number in Driver Files'
change_ver(__FILE__)  #give  the path of this folder


puts 'Updates Complete'