require 'win32ole'
require 'pathname'
require 'fileutils'

# http://stackoverflow.com/questions/1006923/automating-office-via-windows-service-on-server-2008/1680214#1680214
# http://stackoverflow.com/questions/9213460/opening-powerpoint-presentations-in-ruby-via-win32ole
module PresentationConverter

  class Service
    PpSaveAsFileType = {
      :ppSaveAsPresentation => 1,
      :ppSaveAsPowerPoint7 => 2,
      :ppSaveAsPowerPoint4 => 3,
      :ppSaveAsPowerPoint3 => 4,
      :ppSaveAsTemplate => 5,
      :ppSaveAsRTF => 6,
      :ppSaveAsShow => 7,
      :ppSaveAsAddIn => 8,
      :ppSaveAsPowerPoint4FarEast => 10,
      :ppSaveAsDefault => 11,
      :ppSaveAsHTML => 12,
      :ppSaveAsHTMLv3 => 13,
      :ppSaveAsHTMLDual => 14,
      :ppSaveAsMetaFile => 15,
      :ppSaveAsGIF => 16,
      :ppSaveAsJPG => 17,
      :ppSaveAsPNG => 18,
      :ppSaveAsBMP => 19,
      :ppSaveAsOpenXMLPresentation => 24,
      :ppSaveAsOpenXMLPresentationMacroEnabled => 25,
      :ppSaveAsOpenXMLShow => 28,
      :ppSaveAsOpenXMLShowMacroEnabled => 29,
      :ppSaveAsOpenXMLTemplate => 26,
      :ppSaveAsOpenXMLTemplateMacroEnabled => 27
    }

    def initialize(options = {})
      @options = {:verbose => true}.merge(options)
    end

    def verbose(str)
      puts str if @options[:verbose]
    end

    def convert(input_file)
      input_file_path = Pathname.new(input_file).realpath
      basename = File.basename(input_file_path, '.*')
      extname = input_file_path.extname
      output_path = Pathname.new(mktmpdir)

      @pp = WIN32OLE.new('PowerPoint.Application')
      fso = WIN32OLE.new("Scripting.FileSystemObject")

      raise "#{input_file_path.to_s} not readable by current process" unless File.readable? input_file_path.to_s

      verbose "Ensuring destination path"
      FileUtils.mkdir_p output_path

      verbose "Converting full presentation"
      [['.pptx', :ppSaveAsOpenXMLPresentation], ['.ppt', :ppSaveAsPresentation]].each do |pair|
        file_name = output_path + ('full' + pair.first)
        if extname == pair.first
          verbose "Copying #{pair.first}"
          FileUtils.copy input_file_path, file_name
        else
          fname = fso.GetAbsolutePathName(input_file_path.to_s)
          verbose "Converting #{pair.first} from #{fname}"
          presentation = @pp.Presentations.Open(fname)
          presentation.SaveAs(file_name.to_s, PpSaveAsFileType[pair.last], false)
          presentation.Close()
        end
      end

      verbose "Converting individual slides"
      num = 1
      loop do
        puts "Processing slide #{num}"
        presentation = @pp.Presentations.Open(input_file_path.to_s)
        slides = presentation.Slides
        is_last = slides.Count == num
        if slides.Count <= num
          slides.Item(num).MoveTo(1)
          while slides.Count > 1 do
            slides.Item(2).Delete()
          end
          file_name = output_path + ("slide-" + num.to_s.rjust(3, '0'))
          slides.Item(1).Export(file_name.to_s + ".png", "PNG")
          presentation.SaveAs(file_name.to_s, PpSaveAsFileType[:ppSaveAsPresentation], false)
          presentation.SaveAs(file_name.to_s, PpSaveAsFileType[:ppSaveAsOpenXMLPresentation], false)
        end
        presentation.Close()
        num = num + 1
        break if is_last
      end

      output_path
    ensure
      @pp.Quit() if @pp
      @pp = nil
    end

    def mktmpdir
      path = File.join(File.dirname(__FILE__), "tmp", (0...10).map{ ('a'..'z').to_a[rand(26)] }.join)
      FileUtils.mkdir_p path
      FileUtils.chmod 0755, path
      path
    end
  end
end

if __FILE__ == $0
  raise "Missing argument 1" unless ARGV.first
  outpath = PresentationConverter::Service.new.convert(ARGV.first)
  puts outpath.to_s
end
