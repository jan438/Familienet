import turtle
t = turtle.Turtle()
canvas = t.getscreen().getcanvas()

canvas.postscript(file="PDF/turtle_img.ps")

key = input("Wait")
