=begin rdoc
*Revisions*
  | Change                                               | Name        | Date  |

*Test_Script_Name*
  dns_test_table_info

*Test_Case_Number*
  pending

*Description*
  Validate DNS Test Table information

*Variables*
    s = test start time
    f = test finish time
    roe = row number in controller spreadsheet if called from controller
    excel = nested array containing an instance of excel and script parameters
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

*Methods*
    setup - location: setup module
      - create time stamped controller spreadsheet
      - open IE or attache to existing IE session
      - populate the spreadsheet with web support info if ran independently

    logn_chk - location: setup module
      - navigate with 'click'; script is called from controller; no login req'd
      - navigate with 'click_no_wait'; script called from directly; login req'd

    table_info - location: generic class
       - reads configuration item table information from web page
       - write configuration item table information to spreadsheet

    tear_down_d - location: teardown module
      - writes the finished and elapsed time to the controller spreadsheet
      - reads pass/fail status from the driver spreadsheet
      - save and close the driver spreadsheet
      - write pass/fail status to controller spreadsheet if called by controller
      - write elapsed time to controller spreadsheet if called by controller
      - save controller spreadsheet if called by controller
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
  g.logn_chk(g.dnstest,excel[1])
 
  g.table_info(1,4,2,2,ws)

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

 






