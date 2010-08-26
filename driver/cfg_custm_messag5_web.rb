=begin rdoc
*Revisions*
  | Change                                               | Name        | Date  |

*Test_Script_Name*
  custm_messag

*Test_Case_Number*
  700.160.20.110

*Description*
  Validate the Customize Messaging Information configuration
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

#Launch the Customize Messaging ruby script
#Add library file to the path
$:.unshift File.dirname(__FILE__).chomp('driver')<<'lib' # add library to path
s = Time.now
require 'generic'
require 'watir/process' 

  # Write checkbox value
  # TODO This method needs to move into Generic
 def wrt_ckbox(checkbox,value) # Write checkbox value
    if value == 'set'
      checkbox.set
    else
      checkbox.clear
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
  #Navigate to the 'Configure� tab
  g.config.click
  $ie.maximize  
  #Click the Customize Messaging link on the left side of widow
  #Login if not called from controller
  g.logn_chk(g.custmsg,excel[1])

  # email and sms array for customize messaging check box and text box objects
  email = [g.email_addr,g.email_evtdesc,g.email_name,g.email_cont,g.email_loc,
           g.email_desc,g.email_weblnk,g.email_consol,g.email_consoltime,
           g.email_consolevt]
  sms   = [g.sms_addr,g.sms_evtdesc,g.sms_name,g.sms_cont,g.sms_loc,g.sms_desc,
           g.sms_weblnk,g.sms_consol,g.sms_consoltime,g.sms_consolevt]
  

  row = 1
  while(row <= rows)
    puts "Test step #{row}"
    row +=1 # add 1 to row as execution starts at row 2 of driver spreadsheet
    sleep 3
    Watir::Waiter.wait_until(5) {g.edit.exists?}
    g.edit.click

    # Write customize email and sms items
    flag = ws.Range("u#{row}")['Value']
    puts "The flag is #{flag}"
    
    type = email if flag == 'email'
    type = sms if flag == 'smsl'

    # value is a collection of all inputs for a given row
    value =  (ws.Range("k#{row}:t#{row}")['Value']).flatten!
    x = 0 # initialize data array index iterator

    # items are the customize message checkboxes and text
    type[0..9].each do|item|
      wrt_ckbox(item,value[x]) if x < 8
      item.set value[x].to_s if x > 7 && value[7] == 'set'
      x += 1
    end

   #Is there a popup expected? 
    pop = ws.Range("af#{row}")['Value'].to_s 
    puts "  pop_up value = #{pop}" unless pop == 'no'
    #sleep 1

    #If popup, handle with reset OK or reset Cancel to continue
    if (pop == "res" or pop == "can")
      popup_txt  = g.invChar($ie,pop,nil)
      puts "Pop-Up text is #{popup_txt}"
      ws.Range("bk#{row}")['Value'] = popup_txt
    end

    #If reset Cancel, do not save
    if (pop == "no")
      p'no pop save no wait'
      g.save.click_no_wait
      g.jsClick('Windows Internet Explorer', 'OK')
    end
 
    #read Email and SMS Checkbox values
    sleep 3 # increase from 1 to 3 seconds
    Watir::Waiter.wait_until(5) { g.edit.exists?}
    g.edit.click

    col = "bc"
    if flag == 'email'
      email[0..7].each do|checkbox|
        ws.Range("#{col}#{row}")['Value'] = g.checkbox(checkbox)
        col = col.next
      end
      if g.checkbox(g.email_consol) == 'set'
        ws.Range("bl#{row}")['Value'] = g.email_consoltime.value
        ws.Range("bm#{row}")['Value'] = g.email_consolevt.value
      end
    else
      sms[0..7].each do|checkbox|
        ws.Range("#{col}#{row}")['Value'] = g.checkbox(checkbox)
        col = col.next
      end
      if g.checkbox(g.sms_consol) == 'set'
        ws.Range("bl#{row}")['Value'] = g.sms_consoltime.value
        ws.Range("bm#{row}")['Value'] = g.sms_consolevt.value
      end
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