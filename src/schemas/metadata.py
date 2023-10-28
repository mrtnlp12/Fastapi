def metadataEntity(item) -> dict:
    if item is None:
        return None

    clientes = []
    for cliente in item["clientes"]:
        cliente = dict(cliente)
        clientes.append({"nombre": cliente["nombre"], "fila": cliente["fila"]})

    return {
        "fuenteTop": item["fuenteTop"],
        "concentradoTop": item["concentradoTop"],
        "insertados": item["insertados"],
        "clientes": clientes,
    }
