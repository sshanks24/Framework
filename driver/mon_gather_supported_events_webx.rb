=begin rdoc
*Revisions*
  | Initial File | Scott Shanks | 07/2/10 |

*Test_Script_Name*
  mon_gather_supported_events

*Test_Case_Number*
  N/A

*Description*
  Gather all supported events from http and write to a spreadsheet

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

  g.count_frames
  g.monitor.click
  sleep(2)
  while g.count_images('folderplus.gif') > 0
    g.click_all('folderplus.gif')
  end

  g.populate_links_array
  supported_events = Array.new
  g.links_array[1].each do |link|
    begin
      #We don't want to click links with parenthesis for this test case
      unless link.text =~ /\(\d*\)/ then
        puts "Trying link: #{link.text}"
        $ie.frame(:index, 2).link(:id, link.id).click
        #Wait for the table to finish populating
        sleep(2)
        if link.text =~ /\[\d*\]/ then
          #Need to figure out how to append the link.text to this array
          supported_events << g.get_table_by_title('Supported Events')
        else
          supported_events << g.get_table_by_title('Supported Events')
        end
      end
    rescue => e
      if e.to_s =~ /unknown property or method/ then
        next #Ignore this error and continue the loop - I think it raises an
             #exception because it is trying to access the text of an image...
      else
        puts e.to_s
      end
    end
  end
  supported_events = supported_events.flatten.compact
  supported_events.delete_if{|i| i == ""}
  puts supported_events.inspect
  supported_events.to_spread_sheet(ws,2,g.row_ptr) unless supported_events.nil?

  f = Time.now  #finish time
rescue Exception => e
  f = Time.now  #finish time
  puts" \n\n **********\n\n #{$@ } \n\n #{e} \n\n ***"
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
