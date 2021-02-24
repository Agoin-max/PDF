class Singleton:
    __instance = None

    # 单例模式：保障当前类创建的对象只有1个（懒汉模式）
    def __new__(self, *args, **kwargs):
        # if Singleton.__instance:
        #     return self
        # else:
        #     Singleton.__init__(self, *args, **kwargs)
        #     return self
        if not Singleton.__instance:
            Singleton.__init__(self, *args, **kwargs)
        return self

    def __init__(self, data=""):
        self.data = data
        Singleton.__instance = self


# 先执行__new__再执行__init__
s1 = Singleton("数据1")
s2 = Singleton("数据2")
print(s1.data)
print(s2.data)
