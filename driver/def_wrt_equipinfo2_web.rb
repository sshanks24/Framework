=begin rdoc
*Revisions*
  | Change                                               | Name        | Date  |

*Test_Script_Name*
  def_wrt_euipinfo1.rb

*Test_Case_Number*
  700.010.20.110

*Description*
  Validate the Agent Information configuration
     - Edge
     - Boundary
     - Inval character in Name, Contact, Location, and Description

*Variable_Definitions*
    s = test start time
    f = test finish time
    e = test elapsed time
    roe = row number in controller spreadsheet
    excel = nested array that contains an instance of excel and driver parameters
    ss = spreadsheet
    wb = workbook
    ws = worksheet
    dvr_ss = driver spreadsheet
    rows = number of rows in spreadsheet to execute
    site = url/ip address of card being tested
    name = user name for login
    pswd = password for login

=end

$:.unshift File.dirname(__FILE__).chomp('driver')<<'lib' # add library to path
s = Time.now
require 'generic' 

begin 
  puts" \n Executing: #{(__FILE__)}\n\n" # print current filename
  g = Generic.new
  roe = ARGV[1].to_i
  excel = g.setup(__FILE__)
  wb,ws = excel[0][1,2]
  rows = excel[1][1] 

  $ie.speed = :zippy
  g.config.click 
  $ie.maximize 
  g.logn_chk(g.equipinfo,excel[1])
  
  row = 1
  while(row <= rows)
    unless row == 1
      g.config.click
      g.equipinfo.click
    end
    puts " Executing -  Test step #{row}"
    row +=1 # add 1 to row as execution starts at drvr_ss row 2
    #write Agent - Name, Contact, Location, and Description
    edit_button = g.edit.exists?
    if (edit_button == true)
      g.edit.click
    else 
      puts "Edit button not found"
    end
    g.name.set(ws.Range("k#{row}")['Value'])
    g.cont.set(ws.Range("l#{row}")['Value'])
    g.loc.set(ws.Range("m#{row}")['Value'])
    g.desc.set(ws.Range("n#{row}")['Value'])
        
end

  f = Time.now  #finish time
rescue Exception => e
  f = Time.now  #finish time 
  puts" \n\n **********\n\n #{$@ } \n\n #{e} \n\n ***"
  error_present=$@.to_s

ensure #this section is executed even if script goes in error
  if(error_present == nil)
    # If roe > 0, script is called from controller
    # If roe = 0, script is being ran independently
    g.tear_down_d(excel[0],s,f,roe)
    if roe == 0
      $ie.close
    end
  else 
    puts" There were errors in the script"
    status = "script in error"
    wb.save
    wb.close
  end
end