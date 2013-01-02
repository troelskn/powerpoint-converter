require 'net/http'
require 'uri'
require 'json'
require 'base64'
require 'fileutils'

module PresentationConverter
  class ClientController

    def initialize(uri, converter)
      @uri = URI.parse(uri)
      @converter = converter
    end

    def process_next
      http = Net::HTTP.new(@uri.host, @uri.port)
      request = Net::HTTP::Post.new(@uri.request_uri)
      request["Content-Length"] = "0"
      response = http.request(request)
      if response.code.to_i == 200
        result = JSON.parse(response.body)
        tmp_dir = mktmpdir
        file_name = File.join(tmp_dir, result['file_name'])
        File.open(file_name, 'wb') do |f|
          f.write(Base64.decode64(result['file_content']))
        end
        output = @converter.convert(file_name)
        post_body = []
        boundary = (0...50).map{ ('a'..'z').to_a[rand(26)] }.join
        Dir.glob((output + '*').to_s) do |file|
          post_body << "--#{boundary}\r\n"
          post_body << "Content-Disposition: form-data; name=\"datafile[#{File.basename(file)}]\"; filename=\"#{File.basename(file)}\"\r\n"
          post_body << "Content-Type: applicaton/octet-stream\r\n"
          post_body << "\r\n"
          post_body << File.read(file)
          post_body << "\r\n"
        end
        post_body << "--#{boundary}--\r\n"
        request = Net::HTTP::Put.new(result['uri'])
        request.body = post_body.join
        request["Content-Type"] = "multipart/form-data, boundary=#{boundary}"
        http.request(request)
      end
    end

    private
    def mktmpdir
      path = File.join(File.dirname(__FILE__), "tmp", (0...10).map{ ('a'..'z').to_a[rand(26)] }.join)
      FileUtils.mkdir_p path
      FileUtils.chmod 0755, path
      path
    end
  end
end

if __FILE__ == $0
  $:.unshift File.dirname(__FILE__)
  require 'service'
  raise "Missing argument 1" unless ARGV.first
  controller = PresentationConverter::ClientController.new(ARGV.first, PresentationConverter::Service.new)
  controller.process_next
end