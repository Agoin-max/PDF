def variton_item(func):
    def wrapper(*args, **kwargs):
        print("验证身份正确")
        return func(*args, **kwargs)

    return wrapper


@variton_item
def func01():
    print("func01执行成功")


def func02():
    print("func02执行成功")


func01()
