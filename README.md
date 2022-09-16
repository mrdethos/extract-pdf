# extract-pdf

Um programa em python que extrai o texto de um .pdf e cria um arquivo word com o conteúdo.
A extração é feita utilizando a biblioteca PyMuPDF, e caso o arquivo possua alguma imagem, a biblioteca pytesseract é utilizada para OCR.
A criação do arquivo word é feita através da biblioteca docx.
