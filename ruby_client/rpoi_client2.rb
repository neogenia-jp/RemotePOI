this_dir = File.expand_path(File.dirname(__FILE__))
lib_dir = File.join(this_dir, 'lib')
$LOAD_PATH.unshift(lib_dir) unless $LOAD_PATH.include?(lib_dir)

require 'rpoi_services_pb'
require_relative 'interceptors'

def create_interceptors
  [
    Interceptors::PerformanceLoggingInterceptor.new,
    Interceptors::SessionSupportInterceptor.new,
  ]
end

# example for multi session
def main(host, output_path)
  stub = Rpoi::RemotePOI::Stub.new(host, :this_channel_is_insecure, interceptors: create_interceptors)
  stub2 = Rpoi::RemotePOI::Stub.new(host, :this_channel_is_insecure, interceptors: create_interceptors)
  stub.use_template_file(Google::Protobuf::StringValue.new value: '/mnt/est_template.xls')
  stub2.use_template_file(Google::Protobuf::StringValue.new value: '/mnt/est_template.xls')
  stub.set_cell_value(Rpoi::CellAddressWithValue.new row: 10, col: 1, value: Rpoi::CellValue.new(valueType: Rpoi::CellValueTypes::String, string_value: 'SDA-36'))
  stub2.set_cell_value(Rpoi::CellAddressWithValue.new row: 10, col: 1, value: Rpoi::CellValue.new(valueType: Rpoi::CellValueTypes::String, string_value: 'SDA-27'))
  stub.set_cell_value(Rpoi::CellAddressWithValue.new row: 10, col: 3, value: Rpoi::CellValue.new(valueType: Rpoi::CellValueTypes::Numeric, numeric_value: 20))
  stub2.set_cell_value(Rpoi::CellAddressWithValue.new row: 10, col: 3, value: Rpoi::CellValue.new(valueType: Rpoi::CellValueTypes::Numeric, numeric_value: 20))
  stub.set_cell_value(Rpoi::CellAddressWithValue.new row: 11, col: 1, value: Rpoi::CellValue.new(valueType: Rpoi::CellValueTypes::String, string_value: 'SDA-27'))
  stub2.set_cell_value(Rpoi::CellAddressWithValue.new row: 11, col: 1, value: Rpoi::CellValue.new(valueType: Rpoi::CellValueTypes::String, string_value: 'SDA-36'))
  stub.set_cell_value(Rpoi::CellAddressWithValue.new row: 11, col: 3, value: Rpoi::CellValue.new(valueType: Rpoi::CellValueTypes::Numeric, numeric_value: 40))
  stub2.set_cell_value(Rpoi::CellAddressWithValue.new row: 11, col: 3, value: Rpoi::CellValue.new(valueType: Rpoi::CellValueTypes::Numeric, numeric_value: 40))
  stub.set_cell_value(Rpoi::CellAddressWithValue.new row: 12, col: 1, value: Rpoi::CellValue.new(valueType: Rpoi::CellValueTypes::String, string_value: 'SDB-40'))
  stub2.set_cell_value(Rpoi::CellAddressWithValue.new row: 12, col: 1, value: Rpoi::CellValue.new(valueType: Rpoi::CellValueTypes::String, string_value: 'SDC-18'))
  stub.set_cell_value(Rpoi::CellAddressWithValue.new row: 12, col: 3, value: Rpoi::CellValue.new(valueType: Rpoi::CellValueTypes::Numeric, numeric_value: 30))
  stub2.set_cell_value(Rpoi::CellAddressWithValue.new row: 12, col: 3, value: Rpoi::CellValue.new(valueType: Rpoi::CellValueTypes::Numeric, numeric_value: 30))
  stub.set_cell_value(Rpoi::CellAddressWithValue.new row: 13, col: 1, value: Rpoi::CellValue.new(valueType: Rpoi::CellValueTypes::String, string_value: 'SDC-18'))
  stub2.set_cell_value(Rpoi::CellAddressWithValue.new row: 13, col: 1, value: Rpoi::CellValue.new(valueType: Rpoi::CellValueTypes::String, string_value: 'SDB-40'))
  stub.set_cell_value(Rpoi::CellAddressWithValue.new row: 13, col: 3, value: Rpoi::CellValue.new(valueType: Rpoi::CellValueTypes::Numeric, numeric_value: 15))
  stub2.set_cell_value(Rpoi::CellAddressWithValue.new row: 13, col: 3, value: Rpoi::CellValue.new(valueType: Rpoi::CellValueTypes::Numeric, numeric_value: 15))

  stub.set_cell_value(Rpoi::CellAddressWithValue.new row: 14, col: 1, value: Rpoi::CellValue.new(valueType: Rpoi::CellValueTypes::String, string_value: 'SDA-36'))
  stub.set_cell_value(Rpoi::CellAddressWithValue.new row: 14, col: 3, value: Rpoi::CellValue.new(valueType: Rpoi::CellValueTypes::Numeric, numeric_value: 650))

  bytes = stub.download(Google::Protobuf::Empty.new).value
  bytes2= stub2.download(Google::Protobuf::Empty.new).value

  num_of_written = IO.binwrite(output_path, bytes)
  puts "OK! #{num_of_written} bytes written to '#{output_path}'!"

  path2 = insert_suffix(output_path,2)
  num_of_written2= IO.binwrite(path2, bytes2)
  puts "OK! #{num_of_written} bytes written to '#{path2}'!"
end

def insert_suffix(path, suffix)
  ext = File.extname(path)
  path.gsub(/#{ext}$/, suffix.to_s) + ext
end

host = ARGV[0]
unless host
  puts "Usage: #{$0} <host>:<port> [output_path]"
  exit 1
end

output_path = ARGV.length >= 2 ? ARGV[1] : '/tmp/aa.xls'

main(host, output_path)

