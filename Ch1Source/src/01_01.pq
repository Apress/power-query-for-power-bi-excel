let
    Source = Csv.Document(File.Contents("C:\Users\Chris\Documents\My Dropbox\Power Query book\First Drafts\Chapter01\Demos\01_01_SimpleSales.csv"),null,null,null,1252),
    FirstRowAsHeader = Table.PromoteHeaders(Source),
    ChangedType = Table.TransformColumnTypes(FirstRowAsHeader,{{"Month", type text}, {" Sales", type number}}),
    KeptFirstRows = Table.FirstN(ChangedType,2)
in
    KeptFirstRows