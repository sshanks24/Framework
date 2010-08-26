=begin rdoc
*Revisions*
  | Initial File | Scott Shanks | 08/09/2010 |

*Test_Script_Name*
  

*Test_Case_Number*
  

*Description*
  Sets and gets values to various read-write oids from a sreadsheet

*Variable_Definitions*

=end

$:.unshift File.dirname(__FILE__).chomp('driver')<<'lib' # add library to path
s = Time.now
require 'generic'

#Columns in the driver spreadsheet
OIDs = 'C'
Min = 'E'
Max = 'F'
Start_Value = 'E'
Out_Low = 'I'
Out_Low_Result = 'J'
Out_High = 'K'
Out_High_Result = 'L'
Valid = 'F'
Expected = 'G'
Valid_Result = 'N'
Data_Type = 'O'

#Constants used for Excel
XLUP = -4162

begin
  puts" \n Executing: (#{__FILE__}).\n\n" # print current filename
  g = Generic.new
  roe = ARGV[1].to_i
  excel = g.setup(__FILE__)
  wb,ws = excel[0][1,2]
  ws = wb.Worksheets('Data')
  $ie.speed = :zippy
  $ie.close

  begin
    results = Array.new
    
    g.snmp_setup(wb)
    ws = wb.Worksheets('Data')
    total_rows = ws.Range("#{OIDs}65536").End(XLUP).Row
    oids = ws.Range("#{OIDs}2:#{OIDs}#{total_rows}")['Value']
    i = 2
    oids.each do |oid|
      results << g.snmp_test_rw_oid(oid,ws.Range("#{Valid}#{i}")['Value'],ws.Range("#{Expected}#{i}")['Value'])
      i+=1
    end
    results.to_spread_sheet(ws,2,2,8)
  end

  f = Time.now  #finish time
rescue Exception => e
  f = Time.now  #finish time
  puts" \n\n **********\n\n #{$@}\n\n #{e}\n\n ***"
  error_present=$@.to_s

ensure #this section is executed even if script goes in error
  results.to_spread_sheet(ws,2,2,8)
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
