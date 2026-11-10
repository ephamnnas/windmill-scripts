def main(name: str = "World"):
    greeting = f"Hello, {name}!"
    return {"greeting": greeting, "length": len(greeting)}