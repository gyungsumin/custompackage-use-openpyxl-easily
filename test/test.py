class Rect:
    def __init__(self, width, height):
        if width < 0 or height < 0:
            raise Exception("너비와 높이는 음수일 수 없습니다")
        self.__width = width
        self.__height = height

    def get_area(self):
        return self.__width * self.__height


rect = Rect(10, 10)

rect.__width = -10  # 아예 영향을 주지 못함

print(rect.get_area())
print(rect.__width)