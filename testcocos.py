import cocos
import pyglet
from cocos.actions import *
from cocos.director import director


class KeyDisplay(cocos.layer.Layer):

    is_event_handler = True

    def __init__(self):
        super(KeyDisplay, self).__init__()
        self.text = cocos.text.Label('Key: ', font_size=18, x=100, y=280)
        self.key_pressd = set()
        self.add(self.text)

    def update_text(self):
        keys = []
        for i in self.key_pressd:
            key_simbol = pyglet.window.key.symbol_string(i)
            keys.append(key_simbol)
        text = 'Keys: ' + ','.join(keys)
        self.text.element.text = text

    def on_key_press(self, key, modifiers):
        self.key_pressd.add(key)
        self.update_text()

    def on_key_release(self, key, modifiers):
        self.key_pressd.remove(key)
        self.update_text()


class MouseDisplay(cocos.layer.Layer):

    is_event_handler = True

    def __init__(self):
        super(MouseDisplay, self).__init__()
        self.posx = 100
        self.posy = 240
        self.text = cocos.text.Label('No mouse event', font_size=18, x=self.posx, y=self.posy)
        self.add(self.text)

    def update_text(self, x, y):
        text = "Mouse at %d,%d" % (x, y)
        self.text.element.text = text

    def on_mouse_motion(self, x, y, dx, dy):
        self.update_text(x, y)

    def on_mouse_drag(self, x, y, dx, dy, buttons, modifiers):
        self.update_text(x, y)

    def on_mouse_press(self, x, y, buttons, modifiers):
        self.posx, self.posy = director.get_virtual_coordinates(x, y)
        self.text.element.x = self.posx
        self.text.element.y = self.posy

if __name__ == '__main__':
    director.init(resizable=True)
    director.run(cocos.scene.Scene(KeyDisplay(), MouseDisplay()))








# class HelloWorld(cocos.layer.Layer):
#     def __init__(self):
#         super(HelloWorld, self).__init__(64, 64, 224, 225)
#
#         label = cocos.text.Label('Hello World',
#                                  font_name='Times new Roman',
#                                  font_size=50,
#                                  anchor_x='center', anchor_y='center')
#         label.position = 320, 240
#         self.add(label, z=0)
#         sprite = cocos.sprite.Sprite('grossini.png')
#         sprite.position = 320, 240
#         sprite.scale = 1
#         self.add(sprite, z=1)
#         scale = ScaleBy(3, duration=2)
#         label.do(Repeat(scale + Reverse(scale)))
#         sprite.do(Repeat(scale + Reverse(scale)))
#
#
#
# if __name__ == '__main__':
#     cocos.director.director.init()
#     hello_layer = HelloWorld()
#     #hello_layer.do(RotateBy(360, duration=10))
#     main_scene = cocos.scene.Scene(hello_layer)
#     cocos.director.director.run(main_scene)
