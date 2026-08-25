import generate, inspect
src = inspect.getsource(generate.get_output_path)
print(src)
