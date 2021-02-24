class Tools:

    @staticmethod
    def find(iterable, func):
        for item in iterable:
            if func():
                yield item

    @staticmethod
    def duplicate_removal(iterable, func):
        """去重"""
        set_list = []
        if len(iterable) > 1:
            for r in iterable:
                isExists = False
                if not set_list:
                    isExists = True
                    set_list.append(r)
                else:
                    for c in set_list:
                        if func(r, c):
                            isExists = True
                if not isExists:
                    set_list.append(r)
        else:
            set_list = iterable
        return set_list

    @staticmethod
    def sort_by_condition(iterable, func):
        """排序"""
        for r in range(len(iterable) - 1):
            for c in range(r, len(iterable)):
                if func(iterable[r], iterable[c]):
                    iterable[r], iterable[c] = iterable[c], iterable[r]
        return iterable


