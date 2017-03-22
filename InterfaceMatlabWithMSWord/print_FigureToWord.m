

function print_FigureToWord(selection, caption, printType)

    switch printType
        case 'Bitmap'
            print -dbitmap
        case 'WithMeta'
            print -dmeta -painters
        case 'tiff'
            print -dtiff
        case 'FromEditMenu'
            editmenufcn(gcf,'EditCopyFigure')
        case 'ScreenCapture'
            screencapture(gcf,[],'clipboard');
    end

    selection.Paste; 
    selection.Style='Heading 1';
    selection.InsertCaption('Figure',['. ' caption]); %Not working
    selection.TypeParagraph;
end
