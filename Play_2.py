import sys
from PyQt5.QtWidgets import QApplication, QGraphicsView, QGraphicsScene, QGraphicsRectItem
from PyQt5.QtCore import Qt, QTimer, QRectF
from PyQt5.QtGui import QBrush, QColor
import random

class SnakeSegment(QGraphicsRectItem):
    def __init__(self, x, y, w, h):
        super().__init__(x, y, w, h)
        self.setBrush(QBrush(QColor(0, 255, 0)))

class Food(QGraphicsRectItem):
    def __init__(self, x, y, w, h):
        super().__init__(x, y, w, h)
        self.setBrush(QBrush(QColor(255, 0, 0)))

class Snake:
    def __init__(self, scene):
        self.scene = scene
        self.segments = []
        self.direction = Qt.Key_Right
        self.create_snake()

    def create_snake(self):
        for i in range(5):
            segment = SnakeSegment(2, i, 1, 1)
            self.segments.append(segment)
            self.scene.addItem(segment)

    def move(self):
        for i in range(len(self.segments)-1, 0, -1):
            self.segments[i].setPos(self.segments[i-1].pos())

        head = self.segments[0]
        if self.direction == Qt.Key_Right:
            head.moveBy(20, 0)
        elif self.direction == Qt.Key_Left:
            head.moveBy(-20, 0)
        elif self.direction == Qt.Key_Up:
            head.moveBy(0, -20)
        elif self.direction == Qt.Key_Down:
            head.moveBy(0, 20)

    def set_direction(self, key):
        self.direction = key

class Game:
    def __init__(self, scene, snake):
        self.scene = scene
        self.snake = snake
        self.food = None
        self.create_food()


    def create_food(self):
        x = random.randint(0, 20) * 20
        y = random.randint(0, 20) * 20
        self.food = Food(x, y, 20, 20)
        self.scene.addItem(self.food)


    def check_collision(self):
        head = self.snake.segments[0]
        if head.collidesWithItem(self.food):
            self.snake.segments.append(SnakeSegment(0, 0, 20, 20))
            self.scene.addItem(self.snake.segments[-1])
            self.scene.removeItem(self.food)
            self.create_food()

class MainWindow(QGraphicsView):
    def __init__(self):
        super().__init__()
        self.scene = QGraphicsScene(0, 0, 400, 400)
        self.setScene(self.scene)
        self.snake = Snake(self.scene)
        self.game = Game(self.scene, self.snake)
        self.timer = QTimer()
        self.timer.timeout.connect(self.update)
        self.timer.start(100)

    def keyPressEvent(self, event):
        self.snake.set_direction(event.key())

    def update(self):
        self.snake.move()
        self.game.check_collision()

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())