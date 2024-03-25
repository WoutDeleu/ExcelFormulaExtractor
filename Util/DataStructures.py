class Queue:
    def __init__(self):
        self.queue = []
    def add(self, item):
        self.queue.append(item)
    def pop(self):
        return self.queue.pop(0)
    def is_empty(self):
        return len(self.queue) == 0
    def size(self):
        return len(self.queue)
    def get_list(self):
        return self.queue
    def __str__(self):
        return str(self.queue)


class Stack:
    def __init__(self):
        self.stack = []
    def add(self, item):
        self.stack.append(item)
    def pop(self):
        return self.stack.pop()
    def is_empty(self):
        return len(self.stack) == 0
    def size(self):
        return len(self.stack)
    def get_list(self):
        return self.stack
    def __str__(self):
        return str(self.stack)