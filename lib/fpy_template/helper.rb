module FpyTemplate
    class Helper

        #to find a col with a header val
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
        
        #duplicate input and return new workbook
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

        #verify cols and raise exception

        private_class_method def self.verify_cols(output, cols)
            cols.each do |val|
                index_of_val = find_col(output, val)
                if index_of_val.nil?
                raise FpyTemplate::MissingColumns, "el archivo debe tener las columnas identifier, username, email, position y area"  
                end
            end
        end
        
        #verify a col has only the values defined in cols
        private_class_method def self.verify_values_of_columns_exist(file,name,cols)
            col_index = find_col(file, name)
            cols = cols.map(&:downcase)
            i = 1
            while i < file.sheet_data.rows.size
                val = file[i][col_index]&.value
                (raise FpyTemplate::PositionNotDefined, "algunos cargos no existen") unless cols.include?(val&.downcase) || val.nil?
                i+=1
            end
        end
            
        # peform trim
        private_class_method def self.clean_data_for_column(file,cols)
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
            
        # make the defined col value downcased
        private_class_method def self.downcase_cols(file,cols)
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

        # to add cels
        private_class_method def self.print(file,row,col,val)
            if file[row][col]&.value.to_s.empty?
                file&.add_cell(row,col,val) 
            else
                file[row][col]&.change_contents("#{file[row][col]&.value}, #{val}") 
            end
        rescue
            file&.add_cell(row,col,val) 
        end

        #creates a col if it doesnt exist and return index after creating or if already present
        private_class_method def self.create_column_and_return_index(file, name)
            col_index=find_col(file, name)
            return col_index if col_index
            i=file[0]&.size
            return nil if !i
            print(file,0,i,name.to_s)
            return i
        end

        # step 1 
        private_class_method def self.create_autoevaluation_and_ponderation_autoevaluation(file)
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

        # find in file in col with name header and having val in cells then return array of indices
        private_class_method def self.indices_of_rows_with_value_in_col(file,name,val)
            col_index = find_col(file, name)
            arr=[]
            file.each_with_index do |row,i|
                if row[col_index]&.value&.downcase==val&.downcase
                arr<<i
                end
            end
            return arr
        end

        # for top down or bottom up, in a file find rows with position values in to_arr loop through 
        # them in each iteration find all rows with position value of that iteration then start a loop through file and 
        #  set update i flag to false for each row if for this row position is in asendentee array then loop through the array 
        # obtained for the rows having current to_arr value and add  cells in header row with +1 and value in header var,
        # if a value already exists move in next col untill empty cell is found and if the current identifier is not equal
        # to the id of row being updated then only add the value and set flag update_i so i could be incremented for next column 
        private_class_method def self.add_asendentees(file, to_arr , asendentees_arr , header)
            to_arr=to_arr.map(&:downcase)
            asendentees_arr=asendentees_arr.map(&:downcase)
            position_index = find_col(file, "position")
            identifier_index = find_col(file, "identifier")

            to_arr.each do |val|
                indices_of_rows_with_this_position = indices_of_rows_with_value_in_col(file,"position", val)
                i=1
                file.each do |row|
                    update_i=false
                    if asendentees_arr.include?(row[position_index]&.value&.downcase)
                        indices_of_rows_with_this_position.each do |val|

                            ascendente_index = find_col(file,"#{header}_#{i}")
                            while !ascendente_index.nil? && !file[val][ascendente_index]&.value.to_s.empty? 
                                i+=1
                                ascendente_index = find_col(file,"#{header}_#{i}")
                            end
                            if file[val][identifier_index]&.value != row[identifier_index]&.value
                                update_i=true
                                if ascendente_index.nil?
                                    ascendente_index =file[0].size
                                    file.add_cell(0,ascendente_index,"#{header}_#{i}")
                                end
                                file.add_cell(val,ascendente_index,row[identifier_index]&.value)
                            end
                        end
                        i+=1 if update_i
                    end
                end
            end
        end

        #to calculate ponderation of step 7
        private_class_method def self.set_ponderation_values(file)
            ponderation_ascendente_index = find_col(file, "ponderation_ascendente")
            ascendente_index = find_col(file, "ascendente_1")
            ponderation_gerenete_index = find_col(file, "ponderation_gerenete")
            gerente_index = find_col(file, "gerente_1")
            ponderation_descendente_index = find_col(file, "ponderation_descendente")
            descendente_index = find_col(file, "descendente_1")
            ponderation_equipo_index = find_col(file, "ponderation_equipo")
            equipo_index = find_col(file, "equipo_1")
            file.each_with_index do |row, i |
                next if i==0
                arr_of_ponderations_to_fill = []
                if !row[ascendente_index]&.value.to_s.empty?
                    arr_of_ponderations_to_fill << ponderation_ascendente_index
                end
                if !row[gerente_index]&.value.to_s.empty?
                    arr_of_ponderations_to_fill << ponderation_gerenete_index
                end
                if !row[descendente_index]&.value.to_s.empty?
                    arr_of_ponderations_to_fill << ponderation_descendente_index
                end
                if !row[equipo_index]&.value.to_s.empty?
                    arr_of_ponderations_to_fill << ponderation_equipo_index
                end
                total = 0 
                arr_of_ponderations_to_fill.each do |val|
                    ponderation = 100/arr_of_ponderations_to_fill.size
                    total+=100/arr_of_ponderations_to_fill.size
                    if total >= 98 && total<100
                        ponderation += 100-total
                    end
                    file.add_cell(i, val, ponderation)
                end
            end
        end

    end
end