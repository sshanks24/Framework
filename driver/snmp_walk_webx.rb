=begin rdoc
*Revisions*
  | Initial File | Scott Shanks | 07/09/2010 |

*Test_Script_Name*
  

*Test_Case_Number*
  

*Description*
  Parses output from a net-snmp snmpwalk command

*Variable_Definitions*

=end

$:.unshift File.dirname(__FILE__).chomp('driver')<<'lib' # add library to path
s = Time.now
require 'generic'

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
    g.snmp_setup(wb)
    g.snmp_walk(g.test_site,g.community_string,'2c','private').to_spread_sheet(ws,4,2)
  end

  f = Time.now  #finish time
rescue Exception => e
  f = Time.now  #finish time
  puts" \n\n **********\n\n #{$@}\n\n #{e}\n\n ***"
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
