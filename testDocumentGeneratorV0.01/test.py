from nicegui import ui

class DataModel:
    content1 = ''
    content2 = ''
    count = 0

    @property
    def result(self):
        print('运行了该函数', self.count)
        return self.content1 + self.content2

dm = DataModel()

label = ui.label('').bind_text_from(dm, 'result')
input = ui.input('演示').bind_value_to(dm, 'content1')
input = ui.input('演示').bind_value_to(dm, 'content2')

ui.run()
