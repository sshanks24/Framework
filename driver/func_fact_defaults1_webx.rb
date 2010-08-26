=begin rdoc
*Revisions*
  | Change                                               | Name        | Date  |

*Test_Script_Name*
  func_fact_defaults1.rb

*Test_Case_Number*
  700.020.10.110

*Description*
  Validate the Factory Default Information configuration
     Select
     - Deselect
     - Reset Ok and Reset Cancel

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

#Launch the Factory Default ruby script
#Add library file to the path
$:.unshift File.dirname(__FILE__).chomp('driver')<<'lib' # add library to path
s = Time.now
require 'generic'
require 'watir/process'

#Defien bootmode function 
def bootmode(g)
	g.netset.click
	if (g.net_bootmode(0).checked? == true)
	mode = "static"  # mode will equate to DHCP, Static, or BootP
	return mode
    end
end	

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
  #Navigate to the ?Configure? tab
  g.config.click 
  $ie.maximize
  #Click the Factory Default link in the on the left side of window
  #Login if not called from controller
  g.logn_chk(g.factdef,excel[1])
  
  row = 1
  while(row <= rows)
    puts "Test step #{row}"
    row +=1 # add 1, execution starts at drvr_ss row 2
    sleep 1
	
    bt_mode = bootmode(g)
	puts "#{bt_mode}"
	g.factdef.click	
	
	#Is there a popup expected? 
    pop = ws.Range("af#{row}")['Value'].to_s 
    puts "  pop_up value = #{pop}" unless pop == 'reset'
    sleep 1
	puts "  pop_up value = #{pop}"
	
   #Click on Reset to Factory Defaults button  	
   if (pop == "fact")
    g.save.click_no_wait
    popup_txt  =  g.jsClick( $ie,"OK",user_inp = nil,"rtxt")
	puts "Pop-Up text is #{popup_txt}"
    ws.Range("bk#{row}")['Value'] = popup_txt
	puts "#{bt_mode}"
     if (bt_mode != "DHCP")
      puts "Manually re-configure IP and press enter to continue..." 
      STDIN.gets
      #enter = gets
     end
   else
    g.save.click_no_wait
    popup_txt  =  g.jsClick( $ie,"Cancel",user_inp = nil,"rtxt")
	puts "Pop-Up text is #{popup_txt}"
    ws.Range("bk#{row}")['Value'] = popup_txt	
   end
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