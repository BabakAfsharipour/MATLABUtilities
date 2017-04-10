
function print_StructToWord(selection,opt)

    for n=1:length(opt)
        Options = convert_Struct2Cells(opt(n));
        Options = arrange_CellsIntoArray(Options,1);
        Options = convert_CellArray2Table(Options);
        Options.Properties.VariableNames{end} = 'Settings';
        
        print_TableToWord(selection,Options);
        selection.TypeParagraph;
    end
    selection.TypeParagraph;
end


function [cellsOut] = convert_Struct2Cells(theStruct)
    
    if isstruct(theStruct)
        fields = fieldnames(theStruct);
        
        for i=1:length(fields)
            tmp = theStruct.(fields{i});
            
            if ischar(tmp) || isnumeric(tmp) % Return the value
                cellsOut{i,2} = tmp;
                cellsOut{i,1} = fields{i};
            end
            
            if iscategorical(tmp)
                cellsOut{i,2} = tmp;
                cellsOut{i,1} = fields{i};
            end
            
            if iscell(tmp)
                for j=1:length(tmp)
                    str_out = [];
                    if isa(tmp{1},'function_handle')
                        str_out = ['{' num2str(j) '}: ' func2str(tmp{j}) char(13);];
                    elseif ischar(tmp{1})
                        str_out = ['{' num2str(j) '}: ' tmp{j} char(13);];
                    end

                    
                end
                cellsOut{i,2} = str_out;
                cellsOut{i,1} = fields{i};
            end
             
%             if iscell(tmp)
%                 cellsOut{i,2} = tmp{1};
%                 cellsOut{i,1} = fields{i};
%             end
%             
%             if isa(tmp,'function_handle')
%                 cellsOut{i,2} = func2str(tmp);
%                 cellsOut{i,1} = fields{i};
%             end
            
            if isstruct(tmp)
                cellsOut{i,1} = fields{i};
                try
                    [cellsOut{i,2}]=convert_Struct2Cells(tmp);
                catch
                    a=1;
                end
            end
        end
    end
   
end


function cellArray = arrange_CellsIntoArray(cells_In,col_in)
    
    col = col_in;
    cellArray = {};
    for n=1:size(cells_In,1)
                
        if isnumeric(cells_In{n,2}) || ischar(cells_In{n,2})
            cellArray{end+1,col} = cells_In{n,1};
            cellArray{end,10}    = cells_In{n,2};
        end
        
        if iscategorical(cells_In{n,2})
            cellArray{end+1,col} = cells_In{n,1};
            cellArray{end,10}    = cells_In{n,2};
        end
        
        if iscell(cells_In{n,2})
            col = col_in+1;
            tmp = arrange_CellsIntoArray(cells_In{n,2},col);
            tmp(1,col-1) = cells_In(n,1);
            [r,~] = size(tmp); 
            cellArray((end+1):(end+r),:) = tmp;   
        end
        col = col_in;
    end
    
end

function table_out = convert_CellArray2Table(level)
    
    table_out = cell2table(level);
    
    killRow = [];
    for n=1:size(table_out,2)
        tmp = cellfun(@isempty,table_out{:,n});
        if sum(tmp)==length(tmp)
           killRow = [killRow;n]; 
        end    
    end
    
    table_out(:,killRow) = [];
end
