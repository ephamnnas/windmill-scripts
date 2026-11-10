def main(name: str = "there"):
    greeting = f"Hello, {name}!"
    return {"greeting": greeting, "length": len(greeting)}