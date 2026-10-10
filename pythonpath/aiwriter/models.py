"""
Tabla de capacidades de los modelos de OpenAI.

Sustituye las comprobaciones sueltas por nombre de modelo: cada entrada asocia
un prefijo de nombre con las capacidades que difieren de las de por defecto.
"""

# Capacidades que se asumen para cualquier modelo que no esté en la tabla
DEFAULT_CAPABILITIES = {
    "temperature": True,
}

# (prefijo del nombre del modelo, capacidades que cambian respecto a las de por defecto).
# Se usa la primera entrada cuyo prefijo coincida, por lo que los prefijos más
# específicos deben ir antes que los más generales.
MODEL_CAPABILITIES = [
    ("gpt-5", {"temperature": False}),
    # GPT-6 (gpt-6-astra, gpt-6-sol, gpt-6.1-sol, gpt-6-luna): según la guía de
    # OpenAI, temperature solo se admite con reasoning_effort "none", y la
    # extensión no lo envía, así que se omite siempre.
    ("gpt-6", {"temperature": False}),
]


def capabilities(model):
    """
    Devuelve las capacidades del modelo indicado.

    :param model: El nombre del modelo (no distingue mayúsculas).
    :return: Un diccionario nuevo con todas las capacidades del modelo.
    """
    name = (model or "").strip().lower()
    result = dict(DEFAULT_CAPABILITIES)
    for prefix, overrides in MODEL_CAPABILITIES:
        if name.startswith(prefix):
            result.update(overrides)
            break
    return result


def supports(model, capability):
    """
    Indica si el modelo admite una capacidad concreta.

    :param model: El nombre del modelo.
    :param capability: La capacidad a consultar (por ejemplo, "temperature").
    :return: True si el modelo la admite.
    """
    return capabilities(model).get(capability, False)
