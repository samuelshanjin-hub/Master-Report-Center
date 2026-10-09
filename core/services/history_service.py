from collections import deque

class HistoryService:
    def __init__(self, max_items=100):
        self.items = deque(maxlen=max_items)
    def add(self, item): self.items.appendleft(item)
    def all(self): return list(self.items)
