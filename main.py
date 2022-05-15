import xlwings as xw
from string import ascii_uppercase
from random import randint
from time import sleep
import sys

#Colors
BROAD_COLORS = (192, 192, 192) #grey
SNAKE_COLORS = (0, 255, 0) #green
SNAKE_HEAD_COLORS = (0, 102, 0) #blue
APPLE = (255, 0, 0) #red
CONTROLS_COLORS = (96, 96, 96) #black
BUTTON_COLORS = (224, 224, 224)

class Snake:
    def __init__(self, speed, width, height):
        # Book setup
        self.book = xw.Book()
        self.sheet = self.book.sheets[0]
        self.sheet.name = "Snake"
        # Board Setup
        self.speed = 1 / speed
        self.width = width
        self.height = height
        self.board_setup()
        # Snake Setup
        self.body = [(int(height / 2),5), (int(height / 2),4), (int(height / 2),3)]
        self.direction = (0,1)
        self.eaten = False
        self.draw_apple()
        #self.display_game_elements()

    def board_setup(self):

        # Add colors
        game_cells = f'B2:{ascii_uppercase[self.width]}{self.height + 1}'
        self.sheet.range(game_cells).color = BROAD_COLORS

        controls_cell = f'B{self.height + 2}:{ascii_uppercase[self.width]}{self.height + 6}'
        self.sheet.range(controls_cell).color = CONTROLS_COLORS

        # Buttons
        self.exit_cell = f'{ascii_uppercase[self.width]}{self.height + 6}'
        self.sheet[self.exit_cell].value = "Exit"

        ### Movement
        self.left_cell = f'C{self.height + 4}'
        self.sheet[self.left_cell].value = "Left"
        self.right_cell = f'E{self.height + 4}'
        self.sheet[self.right_cell].value = "Right"
        self.down_cell = f'D{self.height + 5}'
        self.sheet[self.down_cell].value = "Down"
        self.up_cell = f'D{self.height + 3}'
        self.sheet[self.up_cell].value = "Up"

        ### Stylings
        for button in [self.exit_cell, self.left_cell, self.right_cell, self.up_cell, self.down_cell]:
            self.sheet[button].color = BUTTON_COLORS
            self.sheet[button].font.align = "center"

        # Dimensions Cells
        self.sheet[f'B2:B{self.height + 6}'].row_height = 40             

    def display_game_elements(self):

        # Display apple
        self.sheet[self.apple_pos].color = APPLE


        # Display snake
        for index, cell in enumerate(self.body):
            if index == 0:
                self.sheet[cell].color = SNAKE_HEAD_COLORS
            else:
                self.sheet[cell].color = SNAKE_COLORS

    def draw_apple(self):
        
        # Get a random cell
        row = randint(1, self.height)
        col = randint(1, self.width)

        # Check if snake is on the apple
        while(row, col) in self.body:
            row = randint(1, self.height)
            col = randint(1, self.width)

        self.apple_pos = (row, col)

    def move_snake(self):
        if self.eaten:
            new_body = self.body[:]
            self.eaten = False
        else:
            lost_cell = self.body[-1]
            new_body = self.body[:-1]
            self.sheet[lost_cell].color = BROAD_COLORS
        new_head = self.add_direction(new_body[0], self.direction)
        new_body.insert(0, new_head)
        self.body = new_body

    ### Game Logic
    def add_direction(self, cell, direction):
        row = cell[0] + direction[0]
        col = cell [1] + direction[1]
        return (row, col)

    def check_collision(self):
        if self.body[0] == self.apple_pos:
            self.eaten = True
            self.draw_apple()
        
        head = self.body[0]
        body = self.body[1:]
        if head in body:
            self.game_over()
            sys.exit()


    ###

    def input(self):
        selected_cell = self.book.selection.address.replace('$', '')
        
        if selected_cell == self.right_cell and self.direction != (0, -1):
            self.direction = (0,1)
        elif selected_cell == self.left_cell and self.direction != (0, 1):
            self.direction = (0,-1)
        elif selected_cell == self.down_cell and self.direction != (-1, 0):
            self.direction = (1,0)
        elif selected_cell == self.up_cell and self.direction != (1, 0):
            self.direction = (-1,0)

    def game_over(self):
        self.sheet[self.body[0]].value = "GAME OVER"

    def run(self):
        while True:
            if self.book.selection.address.replace('$', '') == self.exit_cell:
                self.game_over()
                break
            sleep(self.speed)
            self.input()
            self.move_snake()
            self.check_collision()
            self.display_game_elements()


def __main__(speed=3, width=12, height=6):
    print("Make sure you have Microsoft Excel installed")
    snake = Snake(speed, width, height)
    print("Loading snake")
    snake.run()

if __name__ == "__main__":
    __main__()