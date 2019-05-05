
from __future__ import unicode_literals
from threading import Timer
from wxpy import *
import requests
import random
bot = Bot(cache_path="botoo.pkl")


def search():
         
         groups = bot.groups()
         for group in groups:
                  print(group)

 
def send_news():
    try:
        Text=u"这里是Python测试"
        # 你朋友的微信名称，不是备注，也不是微信帐号。
        my_groups = bot.groups().search("阅")[0]
        my_groups.send(Text)
        my_groups.send_file("readme.txt")
        groups = bot.groups().search("夸夸群")[0]
        groups.send(Text)
        groups.send_file("readme.txt")
        group = bot.groups().search("后")[0]
        group.send(Text)
        group.send_file("readme.txt")
        # 为了防止时间太固定，于是决定对其加上随机数，1天为每86400秒
        ran_int = random.randint(0,20)
        t = Timer(20+ran_int,send_news)

        t.start()

        
        
        
    except:
 
        # 你的微信名称，不是微信帐号。
        my_friend = bot.groups().search('后')[0]
        my_friend.send(u"今天消息发送失败了")



        
if __name__ == "__main__":
    search()
    send_news()

