# VBA-Api-s
Funções API's do Windows x64 para office x64 também.
Essa API acrescenta o botão maximizar e minimizar nos formulários, basta importar o arquivo .bas dentro do seu projeto e no evento de inicialização do formulário, chamar a função passando o caption do formulário. Ex:

Sub Userform_Initialize()
  IncreseElements (Me.caption)
end sub
