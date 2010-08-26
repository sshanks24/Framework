=begin rdoc
*Revisions*
  | Change                                               | Name        | Date  |

*Test_Script_Name*
  cfg_web

*Test_Case_Number*
  700.200.20.110

*Description*
  Validate the Web Configuration
     - Positive
     - Negative
    
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

#Launch the Configure Web ruby script
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
  #Click the Configure Web link on the left side of window
  #Login if not called from controller
  g.logn_chk(g.cfgweb,excel[1])
  
  row = 1
  while(row <= rows)
    puts "Test step #{row}"
    row +=1 # add 1 to row as execution starts at drvr_ss row 2
    sleep 5
    Watir::Waiter.wait_until(5) { g.edit.exists?}
    g.edit.click

    # Write Configure Web fields
    web_server = ((ws.Range("k#{row}")['Value']).to_i).to_s
    puts "#{web_server}"
    g.websrvr.select_value(web_server)
    
    if web_server == "2"
      g.httpport.set((ws.Range("l#{row}")['Value']).to_s)
    elsif web_server == "3"
      g.httpsport.set((ws.Range("m#{row}")['Value']).to_s)
    else
      puts "The port is disabled"
    end
    
    unless web_server == "1"
      if ws.Range("n#{row}")['Value'] == 'set' then g.pswdprtct.set else g.pswdprtct.clear end
      if ws.Range("o#{row}")['Value'] == 'set' then g.cfgctrl.set else g.cfgctrl.clear end
      g.refresh.set((ws.Range("p#{row}")['Value']).to_s)
    else
      puts "The Web Server is Disabled"
    end
    
    #Is there a popup expected? 
    pop = ws.Range("af#{row}")['Value'].to_s 
    puts "  pop_up value = #{pop}" unless pop == 'no'
    #sleep 1

    #If popup, handle with reset OK or reset Cancel to continue
    if (pop == "res")
      popup_txt  = g.invChar($ie,pop,nil)
      puts "Pop-Up text is #{popup_txt}"
      ws.Range("bk#{row}")['Value'] = popup_txt
    end
 
    if (pop == "can")
      ws.Range("bk#{row}")['Value'] = g.res_can(pop)
    end

    #If reset Cancel, do not save
    if (pop == "no")
      g.save.click_no_wait
      g.jsClick('Windows Internet Explorer', 'OK')
    end
 
    #read Configure Web field values
    sleep 1
    Watir::Waiter.wait_until(5) { g.edit.exists?}
    g.edit.click
    
    web_mode = g.websrvr.value
    puts "Web Server = #{web_mode}"
    ws.Range("bc#{row}")['Value'] = web_mode
    if web_mode == "2"
      ws.Range("bd#{row}")['Value'] = g.httpport.value
    elsif web_mode == "3"
      ws.Range("be#{row}")['Value'] = g.httpsport.value
    else
      puts"Web server is Disabled"
    end
    
    unless web_mode == "1"
      ws.Range("bf#{row}")['Value'] = g.checkbox(g.pswdprtct)
      ws.Range("bg#{row}")['Value'] = g.checkbox(g.cfgctrl)
      ws.Range("bh#{row}")['Value'] = g.refresh.value
    else
      puts "Web Server is Disabled"
    end
    
    g.save.click_no_wait
    g.jsClick('Windows Internet Explorer', 'OK')
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