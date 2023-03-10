import argparse


class MyParser:
    def __init__(self):
        self.parser = argparse.ArgumentParser()
        self.parser.add_argument(dest='filename', help="filename")
        self.parser.add_argument('-o', dest='toname', help='toname')
