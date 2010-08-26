=begin rdoc
*Revisions*
  | Change                                               | Name        | Date  |

*Test_Script_Name*
  def_rd_smsinfo1

*Test_Case_Number*
  700.150.20.110

*Description*
  Validate the Email Messaging Information configuration
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

#Launch the Configure Email ruby script
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
  #Navigate to the 'Configure’ tab
  g.config.click
  $ie.maximize  
  #Click the Configure SMS link on the left side of window
  #Login if not called from controller
  g.logn_chk(g.sms,excel[1])
  
  row = 1
  while(row <= rows)
    puts "Test step #{row}"
    row +=1 # add 1 to row as execution starts at drvr_ss row 2
    
	#Read SMS from and SMS to values
    ws.Range("bc#{row}")['Value'] = g.sms_from.value
    ws.Range("bd#{row}")['Value'] = g.sms_to.value
	
	 #Read SMS Subject-Event and Subject Custom values
    if g.sms_subjecttype(0).checked? == true
      ws.Range("be#{row}")['Value'] = "0"
    elsif g.sms_subjecttype(1).checked? == true
      ws.Range("be#{row}")['Value'] = "1"
	  puts "Custom Subject is #{g.sms_custsubj.nil?}"
	  if g.sms_custsubj.nil? ==true
	  g.sms_custsubj.set('Test123')
      ws.Range("bf#{row}")['Value'] = g.sms_custsubj.value
	  else
	  ws.Range("bf#{row}")['Value'] = g.sms_custsubj.value
      end
	end
	#Read SMTP Server Value
	puts "SMTP Server is #{g.sms_srvr.nil?}"
    if g.sms_srvr.nil? == true
      g.sms_srvr.set('142.130.2.83')
      ws.Range("bg#{row}")['Value'] = g.sms_srvr.value
    else
      ws.Range("bg#{row}")['Value'] = g.sms_srvr.value
    end
	
    #Read Port value
    ws.Range("bh#{row}")['Value'] = g.sms_port.value
    g.save.click_no_wait
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