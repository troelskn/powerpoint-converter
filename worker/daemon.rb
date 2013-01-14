if ARGV.first == "install"
  require 'win32/service'
  require 'shellwords'
  require 'pathname'
  include Win32

  SERVICE_NAME = "presentation_converter"
  action = if Service.exists? SERVICE_NAME
             puts "Configuring #{SERVICE_NAME}"
             :configure
           else
             puts "Installing #{SERVICE_NAME}"
             :create
           end
  Service.send(action,
               :service_name       => SERVICE_NAME,
               :service_type       => Service::WIN32_OWN_PROCESS,
               :description        => 'Converts ppt documents on demand',
               :start_type         => Service::AUTO_START,
               :error_control      => Service::ERROR_NORMAL,
               :binary_path_name   => "c:/Ruby193/bin/rubyw.exe " + Shellwords.escape(Pathname.new($0).realpath.to_s),
               :load_order_group   => 'Network',
               :dependencies       => ['W32Time','Schedule'],
               :display_name       => 'Presentation Converter'
               )
  Service.start(SERVICE_NAME)
else
  begin
    $:.unshift File.dirname(__FILE__)
    require 'service'
    require 'controller'

    config_filename = File.join(File.dirname(__FILE__), "config.json")
    $config = begin
               JSON.parse(IO.read(config_filename))
             rescue Errno::ENOENT => ex
               {}
             end
    $config = {'log_file' => File.dirname(__FILE__) + "\\daemon.log", 'url' => "http://192.168.1.2:3000/conversions/fetch", 'frequency' => 3}.merge($config)
    $controller = PresentationConverter::ClientController.new($config['url'], PresentationConverter::Service.new, $config['log_http'] ? $stdout : nil)

    $stdout.reopen($config['log_file'], "a")
    $stdout.sync = true
    $stderr.reopen($stdout)

    puts "#{Time.now} - Starting Daemon"

    require 'win32/daemon'
    include Win32

    class PresentationConverterDaemon < Daemon
      def service_main
        while running?
          sleep $config['frequency'].to_i
          begin
            $controller.process_next
          rescue Exception => err
            puts "*** Daemon failure #{Time.now} err=#{err}"
          end
        end
      end

      def service_stop
        exit!
      end
    end
    PresentationConverterDaemon.mainloop
  rescue Exception => err
    File.open($config['log_file'], 'a') { |f| f.puts "*** Daemon failure #{Time.now} err=#{err}" }
    raise
  end
end
