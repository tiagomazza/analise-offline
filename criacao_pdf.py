from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4

cnv = canvas.Canvas('relatorio de vendas.pdf')
cnv.save()




