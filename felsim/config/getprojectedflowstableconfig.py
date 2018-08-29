

def get_projected_flows_table_config():
    return (
        (
            ("A", "Semana"),
            ("B", "Año"),
            ("C", "Flujo"),
            ("D", "Flexibilidad"),
            ("E", "Tipo"),
            ("F", "Rubro"),
            ("G", "Fecha"),
            ("H", "Cuenta"),
            ("I", "Detalle"),
            ("J", "Ingreso"),
            ("K", "Egreso"),
            ("L", "Fecha proyección"),
        ), (
            ("A", "week"),
            ("B", "year"),
            ("C", "flow"),
            ("D", "flexibility"),
            ("E", "type"),
            ("F", "category"),
            ("G", "date"),
            ("H", "account"),
            ("I", "details"),
            ("J", "income"),
            ("K", "expense"),
            ("L", "projected_date")
        )
    )
