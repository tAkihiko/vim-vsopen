let g:vsopen_script = expand('<sfile>:p:r:gs!\\!/!') . '.vbs'
command! VSOpen silent exe "!wscript" g:vsopen_script expand("%:p") getcurpos()[1] getcurpos()[2]
