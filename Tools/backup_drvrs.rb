=begin
#------------------------------------------------------------------------------#
Create a backup folder in the driver directory and copy all driver files to it.
This folder will be used as a backup before further processing is performed on
the files
#------------------------------------------------------------------------------#
=end
require 'win32ole'
require 'fileutils'
require 'pathname'

def backup_file(file)
  driver_path = (file).gsub(/Tools.*/,'driver/').gsub('/','\\') #get the driver path
  Dir.chdir(driver_path)
  backup_path  =( driver_path + "backup\\").gsub('/','\\')

  #create driver\backup folder
  if !File.directory?(backup_path)
    FileUtils.makedirs(backup_path)
  end

  #copy every file from \driver folder to driver\backup folder
  list=Dir.entries(driver_path).delete_if{ |e| e=~ /^\..*/|| e=~/backup/i}
  list.each_index do |x|
    FileUtils.cp "#{list[x]}",backup_path if !File::directory?(list[x])
  end
end

puts 'Creating backup of all driver files in ..\driver\backup\  folder'
param = backup_file(__FILE__)
puts "Backup completed"

