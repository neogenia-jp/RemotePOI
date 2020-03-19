require_relative 'lib/rpoi_services_pb'

def main(host)
  stub = Rpoi::RemotePOI::Stub.new(host, :this_channel_is_insecure)
  stub.use_template_file(Google::Protobuf::StringValue.new value: '/mnt/est_template_20200313.xls')
  stub.set_cell_value(Rpoi::CellAddressWithValue.new row: 10, col: 1, value: Rpoi::CellValue.new(ValueType: Rpoi::CellValueTypes::String, StringValue: 'SDA-36'))
  stub.set_cell_value(Rpoi::CellAddressWithValue.new row: 10, col: 3, value: Rpoi::CellValue.new(ValueType: Rpoi::CellValueTypes::Numeric, NumericValue: 20))
  stub.set_cell_value(Rpoi::CellAddressWithValue.new row: 11, col: 1, value: Rpoi::CellValue.new(ValueType: Rpoi::CellValueTypes::String, StringValue: 'SDA-27'))
  stub.set_cell_value(Rpoi::CellAddressWithValue.new row: 11, col: 3, value: Rpoi::CellValue.new(ValueType: Rpoi::CellValueTypes::Numeric, NumericValue: 40))
  stub.set_cell_value(Rpoi::CellAddressWithValue.new row: 12, col: 1, value: Rpoi::CellValue.new(ValueType: Rpoi::CellValueTypes::String, StringValue: 'SDB-40'))
  stub.set_cell_value(Rpoi::CellAddressWithValue.new row: 12, col: 3, value: Rpoi::CellValue.new(ValueType: Rpoi::CellValueTypes::Numeric, NumericValue: 30))
  stub.set_cell_value(Rpoi::CellAddressWithValue.new row: 13, col: 1, value: Rpoi::CellValue.new(ValueType: Rpoi::CellValueTypes::String, StringValue: 'SDC-18'))
  stub.set_cell_value(Rpoi::CellAddressWithValue.new row: 13, col: 3, value: Rpoi::CellValue.new(ValueType: Rpoi::CellValueTypes::Numeric, NumericValue: 15))
  stub.set_cell_value(Rpoi::CellAddressWithValue.new row: 14, col: 1, value: Rpoi::CellValue.new(ValueType: Rpoi::CellValueTypes::String, StringValue: 'SDA-36'))
  stub.set_cell_value(Rpoi::CellAddressWithValue.new row: 14, col: 3, value: Rpoi::CellValue.new(ValueType: Rpoi::CellValueTypes::Numeric, NumericValue: 650))
  bytes = stub.download(Google::Protobuf::Empty.new).value
  num_of_written = IO.binwrite('/tmp/aa.xls', bytes)
  puts "OK! #{num_of_written} bytes written!"
end

host = ARGV[0]
unless host
  puts "Usage: #{$0} <host>:<port>"
  exit 1
end

main(host)

