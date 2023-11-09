import sys
from PyQt5.QtWidgets import QApplication, QWidget, QGridLayout, QPushButton
from PyQt5.QtCore import Qt
import random

class TicTacToe(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.buttons = []
        grid = QGridLayout()
        self.setLayout(grid)

        for i in range(3):
            self.buttons.append([])
            for j in range(3):
                button = QPushButton(' ')
                button.setFixedSize(100, 100)
                button.clicked.connect(self.buttonClicked)
                grid.addWidget(button, i, j)
                self.buttons[i].append(button)

        self.show()

    def buttonClicked(self):
        sender = self.sender()
        if sender.text() == ' ':
            sender.setText('X')
            self.checkWin('X')
            self.aiMove()

    def aiMove(self):
        empty_cells = [(i, j) for i in range(3) for j in range(3) if self.buttons[i][j].text() == ' ']
        if empty_cells:
            i, j = random.choice(empty_cells)
            self.buttons[i][j].setText('O')
            self.checkWin('O')

    def checkWin(self, player):
        for i in range(3):
            if self.buttons[i][0].text() == self.buttons[i][1].text() == self.buttons[i][2].text() == player:
                self.gameOver(player)
            if self.buttons[0][i].text() == self.buttons[1][i].text() == self.buttons[2][i].text() == player:
                self.gameOver(player)
        if self.buttons[0][0].text() == self.buttons[1][1].text() == self.buttons[2][2].text() == player:
            self.gameOver(player)
        if self.buttons[0][2].text() == self.buttons[1][1].text() == self.buttons[2][0].text() == player:
            self.gameOver(player)

    def gameOver(self, player):
        for i in range(3):
            for j in range(3):
                self.buttons[i][j].setEnabled(False)
        self.setWindowTitle(f'Player {player} wins!')

if __name__ == '__main__':
    app = QApplication(sys.argv)
    game = TicTacToe()
    sys.exit(app.exec_())