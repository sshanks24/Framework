=begin rdoc
*Revisions*
  | Change                                               | Name        | Date  |

*Test_Script_Name*
  cfg_firmware_TFTP

*Test_Case_Number*
  700.040.20.110

*Description*
  Validate the Web Configuration
     - Positive
     - Negative
     - Boundary
    
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

#Launch the Configure Firmware TFTP ruby script
#Add library file to the path
$:.unshift File.dirname(__FILE__).chomp('driver')<<'lib' # add library to path
s = Time.now
require 'generic'
require 'watir/process' 

begin 
  puts" \n Executing: #{(__FILE__)}\n\n" # print current filename
  g = Generic.new
  roe = ARGV[1].to_i
  #Open up a new excel instance and save it with the timestamp.
  #Open the IE browser
  #Collect the support page information and save it in the time stamped spreadsheet.
  excel = g.setup(__FILE__)
  wb,ws = excel[0][1,2]
  rows = excel[1][1] 

  $ie.speed = :zippy
  #Navigate to the 'Configure' tab
  g.config.click
  $ie.maximize  
  #Click the Configure Firmware TFTP link on the left side of window
  #Login if not called from controller
  g.logn_chk(g.updttftp,excel[1])
  
  row = 1
  while(row <= rows)
    
    puts "Test step #{row}"
    row +=1 # add 1 to row as execution starts at drvr_ss row 2
    sleep 5
    Watir::Waiter.wait_until(10) { g.edit.exists?}
    g.edit.click

    # Write Configure Firmware TFTP fields
    g.tftp_srvr.set(ws.Range("k#{row}")['Value'])
    #g.tftp_port.set(((ws.Range("l#{row}")['Value']).to_i).to_s)
    g.tftp_file.set(ws.Range("m#{row}")['Value'])
    
    # if a popup is expected, handle with Reset-OK or Reset-Cancel
    # if no popup is expected, save
    pop = ws.Range("af#{row}")['Value']
    case pop
    when "no" then g.save.click
    when "ok" then g.save.click
    else ws.Range("bk#{row}")['Value'] = g.res_can(pop)
    end
 
    #read Configure Web field values
    sleep 1
    Watir::Waiter.wait_until(10) { g.edit.exists?}
    g.edit.click
    
    ws.Range("bc#{row}")['Value'] = g.tftp_srvr.value
    #ws.Range("bd#{row}")['Value'] = g.tftp_port.value
    ws.Range("be#{row}")['Value'] = g.tftp_file.value
    
    g.save.click
    #g.jsClick('Windows Internet Explorer', 'OK')
    wb.Save
  end

  f = Time.now  #finish time
#Capture error if any in the script  
rescue Exception => e
  f = Time.now  #finish time 
  puts" \n\n **********\n\n #{$@ } \n\n #{e} \n\n ***"
  error_present=$@.to_s

ensure #this section is executed even if script goes in error
  if(error_present == nil)
    # If roe > 0, script is called from controller
    # If roe = 0, script is being ran independently
    #Close and save the spreadsheet and thes web browser.
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