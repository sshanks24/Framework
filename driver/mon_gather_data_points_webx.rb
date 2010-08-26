=begin rdoc
*Revisions*
  | Initial File | Scott Shanks | 06/17/10 |

*Test_Script_Name*
  mon_gather_data_points

*Test_Case_Number*
  multiple

*Description*
  Gathers all datapoints from the monitor tab and write to a spreadsheet.
This script uses the populate_links_array method to build an array of all text
based links on the navigation pane.

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

  #Clean up the links array for the navigation frame
  
  g.links_array[1].each do |link|
    begin
      #We don't want to click links with parenthesis for this test case
      unless link.text =~ /\(\d*\)/ then
        puts "Trying link: #{link.text}"
        $ie.frame(:index, 2).link(:id, link.id).click
        sleep(2) #Wait for the table to finish populating
        g.table_to_ss(3,ws,link.text)
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
