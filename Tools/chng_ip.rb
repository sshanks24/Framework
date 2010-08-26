=begin
#--------------------------------------------------------------------------------#
Change the device ip in the Controller ss and all of the Driver ss selected in the
Controller ss. This script will run from the \Tools\ folder in the dir structure.
 #-------------------------------------------------------------------------------#
=end

dir = File.dirname(__FILE__)
$:.unshift(dir.gsub('Tools','lib'))# add lib to load path

require 'generic'
g = Generic.new

puts "Enter the Controller spreadsheet number: example - 44:"
ss1_no = gets.chomp!

puts "Enter the IP Address (press enter for 1.2.3.4):"
ip = gets.chomp! 
ip = "1.2.3.4" if ip.empty?

ss1 = dir.sub('Tools','controller/') << 'Controller_' << "#{ss1_no}" << '.xls'
puts " Change IP in Controller => #{ss1}"
xls1 = g.new_xls(ss1,1)
rows = xls1[2].Range("b2")['Value']
xls1[2].Range("b3")['Value'] = ip
xls1[1].Save

puts" rows = #{rows}"
row = 1 
while (row <= rows)
  row += 1 # start at row 2
  run_flag = xls1[2].Range("e#{row}")['Value']
  if run_flag == true
    print" Change IP in Driver xls #{row -1} => "
    path = dir.gsub('Tools','driver/') # driver path
    ss2 = path << (xls1[2].Range("j#{row}")['Value'].to_s).sub('rb','xls') #concat path and xls
    puts "#{ss2}"
    xls2 = g.new_xls(ss2,1) # driver spreadsheet (use new instance)
    xls2[2].Range("b3")['Value'] = ip
    xls2[1].save
    xls2[1].close
    xls2[0].quit
    g.conn_act_xls # connect to controller spreadsheet
  end
end
xls1[0].quit
puts 'done'
