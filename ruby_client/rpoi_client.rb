this_dir = File.expand_path(File.dirname(__FILE__))
lib_dir = File.join(this_dir, 'lib')
$LOAD_PATH.unshift(lib_dir) unless $LOAD_PATH.include?(lib_dir)

require 'rpoi_services_pb'

def main(host, output_path)
  stub = Rpoi::RemotePOI::Stub.new(host, :this_channel_is_insecure,
                                   channel_args: {
                                     # https://github.com/grpc/grpc/blob/b0de95507c51b24279c267489891cdbcc250061c/include/grpc/impl/channel_arg_names.h#L41
                                     'grpc.max_receive_message_length': 20*1024*1024  # set max message size to 20MB
                                   })
  stub.use_template_file(Google::Protobuf::StringValue.new value: '/mnt/est_template.xls')
  stub.set_cell_value(Rpoi::CellAddressWithValue.new row: 10, col: 1, value: Rpoi::CellValue.new(valueType: Rpoi::CellValueTypes::String, string_value: 'SDA-36'))
  stub.set_cell_value(Rpoi::CellAddressWithValue.new row: 10, col: 3, value: Rpoi::CellValue.new(valueType: Rpoi::CellValueTypes::Numeric, numeric_value: 20))
  stub.set_cell_value(Rpoi::CellAddressWithValue.new row: 11, col: 1, value: Rpoi::CellValue.new(valueType: Rpoi::CellValueTypes::String, string_value: 'SDA-27'))
  stub.set_cell_value(Rpoi::CellAddressWithValue.new row: 11, col: 3, value: Rpoi::CellValue.new(valueType: Rpoi::CellValueTypes::Numeric, numeric_value: 40))
  stub.set_cell_value(Rpoi::CellAddressWithValue.new row: 12, col: 1, value: Rpoi::CellValue.new(valueType: Rpoi::CellValueTypes::String, string_value: 'SDB-40'))
  stub.set_cell_value(Rpoi::CellAddressWithValue.new row: 12, col: 3, value: Rpoi::CellValue.new(valueType: Rpoi::CellValueTypes::Numeric, numeric_value: 30))
  stub.set_cell_value(Rpoi::CellAddressWithValue.new row: 13, col: 1, value: Rpoi::CellValue.new(valueType: Rpoi::CellValueTypes::String, string_value: 'SDC-18'))
  stub.set_cell_value(Rpoi::CellAddressWithValue.new row: 13, col: 3, value: Rpoi::CellValue.new(valueType: Rpoi::CellValueTypes::Numeric, numeric_value: 15))
  stub.set_cell_value(Rpoi::CellAddressWithValue.new row: 14, col: 1, value: Rpoi::CellValue.new(valueType: Rpoi::CellValueTypes::String, string_value: 'SDA-36'))
  stub.set_cell_value(Rpoi::CellAddressWithValue.new row: 14, col: 3, value: Rpoi::CellValue.new(valueType: Rpoi::CellValueTypes::Numeric, numeric_value: 650))
  bytes = stub.download(Google::Protobuf::Empty.new).value
  num_of_written = IO.binwrite(output_path, bytes)
  puts "OK! #{num_of_written} bytes written to '#{output_path}'!"
end

host = ARGV[0]
unless host
  puts "Usage: #{$0} <host>:<port> [output_path]"
  exit 1
end

output_path = ARGV.length >= 2 ? ARGV[1] : '/tmp/aa.xls'

main(host, output_path)

