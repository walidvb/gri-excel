require 'rack/app'
require './excel_generator'

class App < Rack::App

  desc 'some hello endpoint'
  get '/hello' do
    serve_file ExcelGenerator.new.file_path
  end

end