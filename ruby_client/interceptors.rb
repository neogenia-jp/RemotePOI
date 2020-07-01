require 'grpc'
require 'pry'

module Interceptors
  class PerformanceLoggingInterceptor < GRPC::ClientInterceptor
    def request_response(request: nil, call: nil, method: nil, metadata: nil)
      start_time = Time.now.to_f
      resp = yield
      request_time = Time.now.to_f - start_time
      puts "duration: #{request_time.round(3)} sec. method: #{method}"
      resp
    end
  end

  class SessionSupportInterceptor < GRPC::ClientInterceptor
    HEADER_KEY = 'x-session-id'
    attr_reader :session_id

    def request_response(request: nil, call: nil, method: nil, metadata: nil)
      if metadata && @session_id
        metadata[HEADER_KEY] = @session_id
      end

      resp = yield

      unless @session_id
      resp_headers = call.instance_variable_get(:@wrapped).metadata
      if resp_headers
        @session_id = resp_headers[HEADER_KEY]
        puts "session_id: #{@session_id}"
      end
      end
      resp
    end
  end
end
