# require "fpy_template/version"
require "byebug"
require "rubyXL"
require 'rubyXL/convenience_methods'

module FpyTemplate
  class MissingColumns < StandardError; end
  class PositionNotDefined < StandardError; end
  class Parse 
    private_class_method def self.find_col(file, name)
      i=0
      found=nil
      while i<file[0]&.size.to_i
          cell = file[0][i]
          if cell&.value&.downcase==name&.downcase
              found=i
              break
          end
          i+=1
      end
      return found
    end
    private_class_method def self.duplicate(input)
      output = RubyXL::Workbook.new
      j = 0
      while j < input.worksheets.size
        output.add_worksheet() if j>0
        input[j].each_with_index do |row, i|
          k=0
          while k < row.size
            output[j].add_cell(i,k,row[k]&.value)
            k+=1
          end
        end
        j+=1
      end
      return output
    end

    def self.verify_cols(output, cols)
      cols.each do |val|
        index_of_val = find_col(output, val)
        if index_of_val.nil?
          raise FpyTemplate::MissingColumns, "el archivo debe tener las columnas identifier, username, email, position y area"  
        end
      end
    end
    
    def self.verify_values_of_columns_exist(file,name,cols)
      col_index = find_col(file, name)
      cols = cols.map(&:downcase)
      i = 1
      while i < file.sheet_data.rows.size
        val = file[i][col_index]&.value
        (raise FpyTemplate::PositionNotDefined, "algunos cargos no existen") unless cols.include?(val&.downcase) || val.nil?
        i+=1
      end
    end
    
    def self.clean_data_for_column(file,cols)
      cols_index = cols.map{|x| find_col(file, x)}
      k = 0
      while k < file.sheet_data.rows.size
        cols_index.each do |i|
          cell = file[k][i]
          cell&.change_contents(cell&.value&.strip&.gsub(/\s+/, " ")&.gsub(/[^[:print:]]/,''))
        end
        k+=1
      end
     
    end
    
    def self.downcase_cols(file,cols)
      cols_index = cols.map{|x| find_col(file, x)}
      k = 0
      while k < file.sheet_data.rows.size
        cols_index.each do |i|
          cell = file[k][i]
          cell&.change_contents(cell&.value&.downcase)
        end
        k+=1
      end
     
    end

    private_class_method def self.print(file,row,col,val)
      if file[row][col]&.value.to_s.empty?
          file&.add_cell(row,col,val) 
        else
          file[row][col]&.change_contents("#{file[row][col]&.value}, #{val}") 
        end
      rescue
        file&.add_cell(row,col,val) 
    end

    def self.create_column_and_return_index(file, name)
      col_index=find_col(file, name)
      return col_index if col_index
      i=file[0]&.size
      return nil if !i
      print(file,0,i,name.to_s)
      return i
    end

    def self.create_autoevaluation_and_ponderation_autoevaluation(file)
      autoevaluacion_index = create_column_and_return_index(file,"autoevaluacion")
      ponderation_autoevaluation_index = create_column_and_return_index(file,"ponderation_autoevaluation")
      position_index = find_col(file, "position")
      skip_for = ["gerente general", "gerente de operaciones" , "obra"]
      i = 1
      while i < file.sheet_data.rows.size
        if !skip_for.include?(file[i][position_index]&.value&.downcase) && !file[i][position_index]&.value.to_s.empty?
          print(file,i,autoevaluacion_index,'x')
          print(file,i,ponderation_autoevaluation_index,0)
        end
        i+=1
      end
      
    end
    def self.indices_of_rows_with_value_in_col(file,name,val)
      col_index = find_col(file, name)
      arr=[]
      file.each_with_index do |row,i|
        if row[col_index]&.value==val
          arr<<i
        end
      end
      return arr
    end
    def self.add_asendentees(file, to_arr , asendentees_arr)
      position_index = find_col(file, "position")
      identifier_index = find_col(file, "identifier")
      added_to = []
      to_arr.each do |val|
        indices_of_rows_with_this_position = indices_of_rows_with_value_in_col(file,"position", val)
        i=1
        file.each do |row|
          if asendentees_arr.include?(row[position_index]&.value&.downcase)
            indices_of_rows_with_this_position.each do |val|
              ascendente_index = find_col(file,"ascendente_#{i}")
              if !ascendente_index
                ascendente_index =file[0].size
                file.add_cell(0,ascendente_index,"ascendente_#{i}")
              end
              file.add_cell(val,ascendente_index,row[identifier_index]&.value)
            end
            i+=1
          end
        end
      end
    end
    def self.transform(file_path)
      input = RubyXL::Parser.parse(file_path)
      output = duplicate(input)
 
      verify_cols(output[0], ["identifier", "username", "email", "position", "area"])
      clean_data_for_column(output[0], ["identifier", "username", "email", "position"])
      downcase_cols(output[0], ["identifier", "username", "email"])
      verify_values_of_columns_exist(output[0],"position", ["Gerente General", "Gerente de Operaciones", "Bodeguero" , "Jefe de Obra", "Prevencionista", "Ayudante profesional", "Administrador de obra", "Profesional de obra","Administrativo","Jefe Administrativo","Capataz","Obra"])
      create_autoevaluation_and_ponderation_autoevaluation(output[0])
      create_column_and_return_index(output[0],"ponderation_ascendente")
      add_asendentees(output[0],["Administrador de Obra"],["Jefe de Obra","Ayudante Profesional","Jefe Administrativo","Jefe de Bodega","Bodeguero","Administrativo","Profesional de Obra","Prevencionista","Capataz"].map(&:downcase))
      add_asendentees(output[0],["Jefe de obra","Ayudante profesional","Jefe administrativo","Ayudante profesional","Jefe de bodega","Administrativo","Bodegero","Prevencionista"],["Capataz"].map(&:downcase))

      output.write('output.xlsx')

    # rescue Exception => e

    #   byebug
    end

  end
end
