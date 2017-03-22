
function print_TableToWord(selection,tbl)

    nr_rows_p = size(tbl,1);
    nr_cols_p = size(tbl,2);
    selection.Tables.Add(selection.Range,nr_rows_p+1,nr_cols_p+1,1,1);
    
    %write the data into the table
    for r=1:nr_rows_p
        
        if r==1
            selection.MoveRight;
            for c=1:nr_cols_p
                selection.TypeText(tbl.Properties.VariableNames{c})
                selection.MoveRight;
                continue;
            end
        end
        
        rowName = tbl.Properties.RowNames{r};
        WordText(selection,rowName,'Normal',[0,0]);
        selection.MoveRight;
        
        for c=1:nr_cols_p
            
            %write data into current cell
            tmp = tbl{r,c}; 
            
            if iscell(tmp)
                tmp = tmp{1};
            end
            if isnumeric(tmp)
                tmp = convert_NumericToString(tmp);
            end
            if iscategorical(tmp)
                tmp = convert_CategoricalToString(tmp);
            end
            if isempty(tmp)
                tmp = '';
            end
            WordText(selection,tmp,'Normal',[0,0]);

            if(c==nr_cols_p)
                %we are done, leave the table
                selection.MoveRight;
                selection.MoveRight;
            else%move on to next cell 
                selection.MoveRight;
            end            
        end
    end
    
    selection.TypeParagraph;
end


function stringOut = convert_CategoricalToString(theCategory)
    
    stringOut = [];
        
    for n=1:length(theCategory)
        tmp = char(theCategory(n));
        if n==1
            stringOut = [tmp];
        else
            stringOut = [stringOut, char(13), tmp];
        end
    end
   
end

function stringOut = convert_NumericToString(theNumber)
    
    stringOut = [];
    if length(theNumber)>2
        stringOut = '';
    elseif length(theNumber)==2
        
        for n=1:2
            tmp{n} = num2str(theNumber(n));
            if length(tmp{n})>4
                tmp{n} = sprintf('%0.3f',theNumber(n));
            end
            
        end
        stringOut = [tmp{1}, char(13), tmp{2}];
    else
        stringOut = num2str(theNumber);
        if length(stringOut)>4
            stringOut = sprintf('%0.3f',theNumber);
        end
    end
end

function WordText(selection,text_p,style_p,enters_p,color_p)

    if(enters_p(1))
        selection.TypeParagraph; %enter
    end
	selection.Style = style_p;
    if(nargin == 5)%check to see if color_p is defined
        selection.Font.ColorIndex=color_p;     
    end
    
	selection.TypeText(text_p);
    selection.Font.ColorIndex='wdAuto';%set back to default color
    for k=1:enters_p(2)    
        selection.TypeParagraph; %enter
    end
    
end