import pathlib
text = pathlib.Path('src/main.py').read_text()
start = text.index('class App(CTk):')
end = text.index('if __name__ == \"__main__\":')
pathlib.Path('app_block.txt').write_text(text[start:end])
