require 'docx'
require 'roo'
require 'write_xlsx'
require 'fileutils'
require 'rubyXL'

class DocProcessor
  PLACEHOLDER_MAPPING = {
    '[nom]' => '%%CLIENT_NAME%%',
    '[naam]' => '%%CLIENT_NAME%%'
    # Add more mappings here
  }.freeze

  def initialize(document)
    @document = document
  end

  def process
    case File.extname(@document.document.path)
    when '.doc', '.docx'
      process_word_document
    when '.xls', '.xlsx'
      process_excel_document
    else
      raise "Unsupported file type"
    end
  end

  def render_document
    path = Rails.root.join('public', @document.processed_document_path).to_s
    Rails.logger.info("Opening processed document at path: #{path}")
    if File.exist?(path)
      case File.extname(path)
      when '.doc', '.docx'
        render_word_document(path)
      when '.xls', '.xlsx'
        render_excel_document(path)
      else
        Rails.logger.error("Unsupported file type at path: #{path}")
        ""
      end
    else
      Rails.logger.error("File does not exist at path: #{path}")
      ""
    end
  rescue StandardError => e
    Rails.logger.error("Error opening document: #{e.message}")
    ""
  end

  def read_excel
    file_path = Rails.root.join('public', @document.processed_document_path)
    Rails.logger.info("Opening processed excel at path: #{file_path}")
    xlsx = Roo::Spreadsheet.open(file_path)
    sheet = xlsx.sheet(0)
    data = []

    sheet.each_row_streaming(offset: 1) do |row|
      row_data = row.map(&:value)
      data << row_data
    end

    data
  end

  def excel_to_html
    file_path = @document.processed_document_path
    workbook = RubyXL::Parser.parse(Rails.root.join('public', file_path))
    html = ""

    workbook.worksheets.each do |worksheet|
      html << "<h2>#{worksheet.sheet_name}</h2>"
      html << "<table border='1'>"
      worksheet.each do |row|
        html << "<tr>"
        row && row.cells.each do |cell|
          html << "<td>#{cell && cell.value}</td>"
        end
        html << "</tr>"
      end
      html << "</table>"
    end

    html
  end

  private

  def process_word_document
    template_path = @document.document.path
    processed_file_path = Rails.root.join('public', 'uploads', 'documents', 'processed', "#{@document.id}.docx").to_s

    # Ensure the processed directory exists
    ensure_processed_directory

    # Open the document
    doc = Docx::Document.open(template_path)

    # Open a log file
    log_file_path = Rails.root.join('log', 'document_processing.log')
    File.open(log_file_path, 'a') do |log|
      log.puts "Processing document: #{template_path}"
      log.puts "Processing started at: #{Time.now}"

      # Replace placeholders with actual values
      doc.paragraphs.each_with_index do |p, p_index|
        log.puts "Paragraph #{p_index + 1}:"

        p.each_text_run do |tr|
          original_text = tr.text
          log.puts "  Original text: #{original_text}"

          PLACEHOLDER_MAPPING.each do |placeholder, replacement|
            if original_text.include?(placeholder)
              new_text = original_text.gsub(placeholder, replacement)
              tr.text = new_text
              log.puts "  Replaced '#{placeholder}' with '#{replacement}' in text: #{new_text}"
            end
          end

          log.puts "  Processed text: #{tr.text}"
        end
      end

      log.puts "Processing finished at: #{Time.now}"
    end

    # Save the modified document
    doc.save(processed_file_path)

    # Update the document record with the new path
    @document.update(processed_document_path: processed_file_path.sub("#{Rails.root}/public", ""))
  end



  def process_word_document_old
    template_path = @document.document.path
    processed_file_path = Rails.root.join('public', 'uploads', 'documents', 'processed', "#{@document.id}.docx").to_s

    # Ensure the processed directory exists
    ensure_processed_directory

    # Open the document
    doc = Docx::Document.open(template_path)

    # Replace placeholders with actual values
    doc.paragraphs.each do |p|
      PLACEHOLDER_MAPPING.each do |placeholder, replacement|
        p.each_text_run do |tr|
          # Replace the placeholder with the actual value
          if tr.text.include?(placeholder)
            tr.text = tr.text.gsub(placeholder, replacement)
          end
        end
      end
    end

    # Save the modified document
    doc.save(processed_file_path)

    # Update the document record with the new path
    @document.update(processed_document_path: processed_file_path.sub("#{Rails.root}/public", ""))
  end

  def process_excel_document
    ensure_processed_directory

    # Open the existing workbook
    workbook = RubyXL::Parser.parse(@document.document.path)

    # Iterate through each worksheet
    workbook.worksheets.each do |worksheet|
      worksheet.each_with_index do |row, row_index|
        next unless row

        row.cells.each_with_index do |cell, col_index|
          next unless cell

          value = cell.value.to_s
          PLACEHOLDER_MAPPING.each do |placeholder, replacement|
            value = value.gsub(placeholder, replacement) if value.is_a?(String)
          end

          # Update cell value
          new_cell = worksheet.add_cell(row_index, col_index, value)

          # Retain cell's original format and style
          new_cell.style_index = cell.style_index
        end
      end
    end

    # Save the new workbook
    processed_file_path = Rails.root.join('public', 'uploads', 'documents', 'processed', "#{@document.id}.xlsx")
    workbook.write(processed_file_path.to_s)

    @document.update(processed_document_path: processed_file_path.to_s.sub("#{Rails.root}/public", ""))
  end

  def process_excel_document_original

    ensure_processed_directory

    xlsx = Roo::Spreadsheet.open(@document.document.path)
    processed_file_path = Rails.root.join('public', 'uploads', 'documents', 'processed', "#{@document.id}.xlsx")

    workbook = WriteXLSX.new(processed_file_path.to_s)

    xlsx.sheets.each do |sheet_name|
      worksheet = workbook.add_worksheet(sheet_name)
      xlsx.sheet(sheet_name).each_with_index do |row, index|
        new_row = row.map do |cell|
          value = cell.is_a?(Roo::Excelx::Cell) ? cell.value.to_s : cell.to_s
          PLACEHOLDER_MAPPING.each do |placeholder, replacement|
            value = value.gsub(placeholder, replacement) if value.is_a?(String)
          end
          value
        end
        worksheet.write_row(index,0, new_row)
      end
    end
    workbook.close
    @document.update(processed_document_path: processed_file_path.to_s.sub("#{Rails.root}/public", ""))
  end

  def ensure_processed_directory
    processed_dir = Rails.root.join('public', 'uploads', 'documents', 'processed')
    FileUtils.mkdir_p(processed_dir) unless Dir.exist?(processed_dir)
  end

  def paragraph_to_html(paragraph)
    html_content = paragraph.text_runs.map do |text_run|
      styles = []
      "<span style=\"#{styles.join(' ')}\">#{text_run.text}</span>"
    end.join
    "<p>#{html_content}</p>"
  end

  def render_word_document(path)
    doc = Docx::Document.open(path)
    doc.paragraphs.map { |p| paragraph_to_html(p) }.join("\n")
  end

  def render_excel_document(path)
    xlsx = Roo::Spreadsheet.open(path)
    sheet = xlsx.sheet(0)
    data = sheet.parse(headers: true)
    table_html = "<table>"
    table_html << "<thead><tr>#{data.headers.map { |header| "<th>#{header}</th>" }.join}</tr></thead>"
    table_html << "<tbody>"
    data.each do |row|
      table_html << "<tr>#{row.map { |cell| "<td>#{cell}</td>" }.join}</tr>"
    end
    table_html << "</tbody></table>"
    table_html
  end
end
