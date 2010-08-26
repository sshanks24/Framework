=begin rdoc
*Revisions*
  | Change                                               | Name        | Date  |

*Test_Script_Name*
  custmmsg_table_info

*Test_Case_Number*
  pending

*Description*
  Validate Customize Message Table information

*Variables*
    s = test start time
    f = test finish time
    roe = row number in controller spreadsheet if called from controller
    excel = nested array that contains an instance of excel and script parameters
      excel[0] = excel instance
        ss = spreadsheet
        wb = workbook
        ws = worksheet
      excel[1] = parameters
        dvr_ss = driver spreadsheet
        rows = number of rows in spreadsheet to execute
        site = url/ip address of card being tested
        name = user name for login
        pswd = password for login
    ARGV[1] = command line parameter passed from the controller for 'roe'
    (ARGV.length !=0) script called from controller else running independently
=end

$:.unshift File.dirname(__FILE__).chomp('driver')<<'lib' # add library to path
s = Time.now
require 'generic' 
  
begin
  puts" \n Executing: #{(__FILE__)}\n\n" # show current filename
  g = Generic.new
  roe = ARGV[1].to_i
  excel = g.setup(__FILE__)
  wb,ws = excel[0][1,2]
  
  g.config.click
  g.logn_chk(g.custmsg,excel[1])
 
  g.table_info(1,11,2,2,ws)

rescue Exception => e
  puts" \n\n **********\n\n #{$@ } \n\n #{e} \n\n ***"
  error_present=$@.to_s
ensure
  f = Time.now  
  if(error_present == nil)
    g.tear_down_d(excel[0],s,f,roe)
    if roe == 0
      $ie.close
    end
  else
    puts" Script execution error"
    wb.Save
    wb.close
  end
end

 






