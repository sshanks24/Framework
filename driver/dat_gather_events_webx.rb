=begin rdoc
*Revisions*
  | Initial File | Scott Shanks | 07/08/2010 |

*Test_Script_Name*
  http_datalogs_gather_events

*Test_Case_Number*
  

*Description*
  Gathers all of the SNMP events from the web interface

*Variable_Definitions*

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

  ws = wb.Worksheets('Data')
  
  g.datalogs.click
  sleep(1)
  g.events.click
  events_text = g.det.text
  a = Array.new
  events_text.each_line do |line|
    if line =~ /\w,\d/ #word character - comma - digit character
      a << line.chomp.split(",")
    end
  end
  a.to_spread_sheet(ws,2,2)

  f = Time.now  #finish time
rescue Exception => e
  f = Time.now  #finish time
  puts" \n\n **********\n\n #{$@} \n\n #{e} \n\n ***"
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
