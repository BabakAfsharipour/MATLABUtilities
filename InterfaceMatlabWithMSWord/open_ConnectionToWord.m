
function [serverHandle, selection] = open_ConnectionToWord()

    serverHandle = actxserver('Word.Application');
    serverHandle.Visible = 1;
    serverHandle.Documents.Add;
    selection = serverHandle.Selection;
end