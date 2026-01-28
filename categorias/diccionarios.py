doc_tipo = [
    "CUIT",
    "CUIL",
    "CDI",
    "CI Extranjera",
    "Pasaporte",
    "DNI",
    "Otro",
    ]

opciones_iva = {
    "iva responsable inscripto": 1, 
    "iva responsable no inscripto": 2,
    "iva no responsable": 3, 
    "iva sujeto exento": 4, 
    "consumidor final": 5, 
    "responsable monotributo": 6, 
    "sujeto no categorizado": 7, 
    "proveedor del exterior": 8, 
    "cliente del exterior": 9, 
    "iva liberado - ley 19.640": 10,
    "iva responsable inscripto - agente de percepción": 11, 
    "pequeño contribuyente eventual": 12, 
    "monotributista social": 13,
    "pequeño contribuytente eventual social": 14, 
    "iva no alcanzado": 15
    }

opciones_codigo_concepto = {
    "Productos": 1,
    "Servicios": 2,
    "Productos y Servicios": 3,
}

opciones_codigo_iva = {
    "No Gravado": 1,
    "Exento": 2,
    "0%": 3,
    "2,50%": 9,
    "5%": 8,
    "10,50%": 4,
    "21%": 5,
    "27%": 6
}

calculo_codigo_iva = {
    "No Gravado": 0,
    "Exento": 0,
    "0%": 0,
    "2,50%": 0.025,
    "5%": 0.05,
    "10,50%": 0.105,
    "21%": 0.21,
    "27%": 0.27
}

opciones_codigo_tributos = {
    "Impuestos Nacionales": 1,
    "Impuestos Provinciales": 2,
    "Impuestos Municipales": 3,
    "Impuestos Internos": 4,
    "IIBB Percepción": 5,
    "Percepción de IVA": 6,
    "Percepción de Ingresos Brutos": 7,
    "Percepción de Impuestos Municipales": 8,
    "Otras Percepciones": 9,
    "Percepción por Resolución General AFIP": 10,
    "Tasa de Seguridad e Higiene": 11,
    "Fondo Financiamiento Ley 25413": 12,
    "Percepción de IVA no Categorizado": 13,
    "Impuestos a los combustibles líquidos y al gas natural": 14,
    "Impuestos a los pasajes al exterior": 15,
    "Impuesto a los créditos y débitos": 16,
    "Percepción a los juegos de azar": 17,
    "Impuesto a los servicios de comunicación audiovisual": 18,
    "Otros": 99,
    "Sin otros tributos": ""
}

opciones_tipo_comprobante = [
    "Comprobante", 
    "Item"
    ]

# Definición de columnas
columns_factura_a = {
    "excel": [
        "Tipo Dato", "Fecha", "Periodo Desde", "Periodo Hasta", "Condicion frente al IVA",
        "Tipo Doc", "Cliente", "Documento", "Domicilio", "Cant.", "Descripcion", "$ Unit.",
        "Total", "Condición Venta", "Unidad Medida", "% Bonificación", "Importe Bonificación",
        "Importe Op Ex", "Importe IVA", "Importe Tributos", "Observaciones", "Concepto",
        "Codigo Otros Tributos", "Descripcion Otros Tributos", "Base Imponible otros Tributos", 
        "UnidadMtx", "CodigoMtx", "Codigo Condición IVA"
    ],
    "display": [
        "Tipo de Dato [Comprobante o Producto de un comprobante anterior]",
        "Fecha de emisión [+-10 días del actual, si excede fallará - YYYY-MM-DD]",  
        "Periodo Desde [Periodo de Facturación: YYYY-MM-DD]",
        "Periodo Hasta [Periodo de Facturación: YYYY-MM-DD]",  
        "Condición frente al IVA [IVA del Receptor - En excel va por código]",  
        "Tipo de Documento [Del que recibe la factura]", 
        "Cliente [En el PDF - Nombre o Razón Social]",  
        "Documento [Número sin guiones ni puntos. Ejemplo CUIT: 30423210137]",  
        "Domicilio [Lugar de residencia del receptor]",  
        "Cantidad [Del producto facturado]", 
        "Descripción [Del producto facturado]", 
        "Precio Unitario [Del producto facturado]",
        "Importe Total [Se calcula: Cantidad x Precio Unitario]",  
        "Condición de Venta [En el PDF de la factura]", 
        "Unidad de Medida [En el PDF de la factura]", 
        "Porcentaje de Bonificación [En el PDF - Ejemplo: 10%]",  
        "Importe Bonificación [Importe bonificado o descuento]",  
        "Importe Operaciones Exentas [Monto de operaciones exentas]",  
        "Importe IVA [Calculado con Total y Código Condición IVA]",  
        "Importe Tributos [Importe de Otros Tributos]",
        "Observaciones [Campo de observaciones - Dejar vacío si no hay]",  
        "Concepto [Código de concepto - 1, 2 o 3]",  
        "Codigo Otros Tributos [Código del tributo aplicado]",  
        "Descripción Otros Tributos [Nombre del tributo aplicado]",  
        "Base Imponible Otros Tributos [Importe gravado del tributo]",  
        "UnidadMtx [Unidades según la matriz de datos]",  
        "CodigoMtx [Código del producto (GTIN-8, GTIN-12, GTIN-13, GTIN-14)]",  
        "Codigo Condicion IVA [Identificador numérico del tipo de IVA]"  
    ]
}

# Definición de columnas
columns_factura_a_credito = {
    "excel": [
        "Tipo Dato", "Punto Venta C. Original", "Número Comprobante Original", "Condicion frente al IVA", "Fecha",
        "Fecha servicio desde", "Fecha servicio hasta", "Fecha vencimiento pago", "Tipo Doc", "Cliente",
        "Documento", "Domicilio", "UnidadMtx", "codigoMtx", "Descripcion", "Cantidad", "Precio Unitario",
        "Importe Total", "Codigo Condicion IVA", "Importe IVA", "Concepto", "Motivo Nota",
        "Importe Otros Tributos", "Codigo Otros Tributos", "Descripcion Otros Tributos", "Base Imponible otros Tributos"
    ],
    "display": [
        "Tipo de Dato [Comprobante o Producto de un comprobante anterior]", 
        "Punto Venta C. Original [Número de punto de venta del comprobante original]", 
        "Número Comprobante Original [Número del comprobante original asociado]", 
        "Condición frente al IVA [IVA del Receptor - En excel va por código numérico]", 
        "Fecha de emisión [+-10 días del actual, si excede fallará - YYYY-MM-DD]", 
        "Fecha servicio desde [Fecha de inicio del período facturado - YYYY-MM-DD]", 
        "Fecha servicio hasta [Fecha de fin del período facturado - YYYY-MM-DD]", 
        "Fecha vencimiento pago [Fecha de vencimiento del pago de la factura - YYYY-MM-DD]", 
        "Tipo de Documento [Del que recibe la factura]", 
        "Cliente [Nombre o Razón Social]", 
        "Documento [Número sin guiones ni puntos. Ejemplo CUIT: 30423210137]", 
        "Domicilio [Lugar de residencia del que recibe la factura]",  
        "UnidadMtx [Unidades según la matriz de datos]", 
        "CodigoMtx [Código del producto (GTIN-8, GTIN-12, GTIN-13, GTIN-14)]", 
        "Descripcion [Del producto facturado]", 
        "Cantidad [Del producto facturado]", 
        "Precio Unitario [Del producto facturado]", 
        "Importe Total [Se calcula: Cantidad x Precio Unitario]", 
        "Codigo Condicion IVA [Identificador numérico del tipo de IVA]", 
        "Importe IVA [Calculado con Importe Total y Código Condición IVA]", 
        "Concepto [Código de concepto - 1, 2 o 3]", 
        "Motivo Nota [Motivo de la nota de crédito o débito, si aplica]", 
        "Importe Otros Tributos [Importe de tributos adicionales aplicados]", 
        "Codigo Otros Tributos [Código del tributo aplicado]", 
        "Descripcion Otros Tributos [Nombre del tributo aplicado]", 
        "Base Imponible otros Tributos [Importe gravado del tributo]"
    ]
}

# Definición de columnas
columns_factura_a_debito = {
    "excel": [
        "Tipo Dato", "Punto Venta C. Original", "Número Comprobante Original", "Condicion frente al IVA", "Fecha",
        "Fecha servicio desde", "Fecha servicio hasta", "Fecha vencimiento pago", "Tipo Doc", "Cliente",
        "Documento", "Domicilio", "UnidadMtx", "codigoMtx", "Descripcion", "Cantidad", "Precio Unitario",
        "Importe Total", "Codigo Condicion IVA", "Importe IVA", "Concepto", "Motivo Nota",
        "Importe Otros Tributos", "Codigo Otros Tributos", "Descripcion Otros Tributos", "Base Imponible otros Tributos"
    ],
    "display": [
        "Tipo de Dato [Comprobante o Producto de un comprobante anterior]", 
        "Punto Venta C. Original [Número de punto de venta del comprobante original]", 
        "Número Comprobante Original [Número del comprobante original asociado]", 
        "Condición frente al IVA [IVA del Receptor - En excel va por código numérico]", 
        "Fecha de emisión [+-10 días del actual, si excede fallará - YYYY-MM-DD]", 
        "Fecha servicio desde [Fecha de inicio del período facturado - YYYY-MM-DD]", 
        "Fecha servicio hasta [Fecha de fin del período facturado - YYYY-MM-DD]", 
        "Fecha vencimiento pago [Fecha de vencimiento del pago de la factura - YYYY-MM-DD]", 
        "Tipo de Documento [Del que recibe la factura]", 
        "Cliente [Nombre o Razón Social]", 
        "Documento [Número sin guiones ni puntos. Ejemplo CUIT: 30423210137]", 
        "Domicilio [Lugar de residencia del que recibe la factura]",  
        "UnidadMtx [Unidades según la matriz de datos]", 
        "CodigoMtx [Código del producto (GTIN-8, GTIN-12, GTIN-13, GTIN-14)]", 
        "Descripcion [Del producto facturado]", 
        "Cantidad [Del producto facturado]", 
        "Precio Unitario [Del producto facturado]", 
        "Importe Total [Se calcula: Cantidad x Precio Unitario]", 
        "Codigo Condicion IVA [Identificador numérico del tipo de IVA]", 
        "Importe IVA [Calculado con Importe Total y Código Condición IVA]", 
        "Concepto [Código de concepto - 1, 2 o 3]", 
        "Motivo Nota [Motivo de la nota de crédito o débito, si aplica]", 
        "Importe Otros Tributos [Importe de tributos adicionales aplicados]", 
        "Codigo Otros Tributos [Código del tributo aplicado]", 
        "Descripcion Otros Tributos [Nombre del tributo aplicado]", 
        "Base Imponible otros Tributos [Importe gravado del tributo]"
    ]
}

# Definición de columnas
columns_factura_b = {
    "excel": [
        "Tipo Dato", "Fecha", "Periodo Desde", "Periodo Hasta", "Condicion frente al IVA",
        "Tipo Doc", "Cliente", "Documento", "Domicilio", "Cant.", "Descripcion", "$ Unit.",
        "Total", "Condición Venta", "Unidad Medida", "% Bonificación", "Importe Bonificación",
        "Importe Op Ex", "Importe IVA", "Importe Tributos", "Observaciones", "Concepto",
        "Codigo Otros Tributos", "Descripcion Otros Tributos", "Base Imponible otros Tributos", 
        "UnidadMtx", "CodigoMtx", "Codigo Condición IVA"
    ],
    "display": [
        "Tipo de Dato [Comprobante o Producto de un comprobante anterior]",
        "Fecha de emisión [+-10 días del actual, si excede fallará - YYYY-MM-DD]",  
        "Periodo Desde [Periodo de Facturación: YYYY-MM-DD]",
        "Periodo Hasta [Periodo de Facturación: YYYY-MM-DD]",  
        "Condición frente al IVA [IVA del Receptor - En excel va por código]",  
        "Tipo de Documento [Del que recibe la factura]", 
        "Cliente [En el PDF - Nombre o Razón Social]",  
        "Documento [Número sin guiones ni puntos. Ejemplo CUIT: 30423210137]",  
        "Domicilio [Lugar de residencia del receptor]",  
        "Cantidad [Del producto facturado]", 
        "Descripción [Del producto facturado]", 
        "Precio Unitario [Del producto facturado]",
        "Importe Total [Se calcula: Cantidad x Precio Unitario]",  
        "Condición de Venta [En el PDF de la factura]", 
        "Unidad de Medida [En el PDF de la factura]", 
        "Porcentaje de Bonificación [En el PDF - Ejemplo: 10%]",  
        "Importe Bonificación [Importe bonificado o descuento]",  
        "Importe Operaciones Exentas [Monto de operaciones exentas]",  
        "Importe IVA [Calculado con Total y Código Condición IVA]",  
        "Importe Tributos [Importe de Otros Tributos]",
        "Observaciones [Campo de observaciones - Dejar vacío si no hay]",  
        "Concepto [Código de concepto - 1, 2 o 3]",  
        "Codigo Otros Tributos [Código del tributo aplicado]",  
        "Descripción Otros Tributos [Nombre del tributo aplicado]",  
        "Base Imponible Otros Tributos [Importe gravado del tributo]",  
        "UnidadMtx [Unidades según la matriz de datos]",  
        "CodigoMtx [Código del producto (GTIN-8, GTIN-12, GTIN-13, GTIN-14)]",  
        "Codigo Condicion IVA [Identificador numérico del tipo de IVA]"  
    ]
}

# Definición de columnas
columns_factura_b_credito = {
    "excel": [
        "Tipo Dato", "Punto Venta C. Original", "Número Comprobante Original", "Condicion frente al IVA", "Fecha",
        "Fecha servicio desde", "Fecha servicio hasta", "Fecha vencimiento pago", "Tipo Doc", "Cliente",
        "Documento", "Domicilio", "UnidadMtx", "codigoMtx", "Descripcion", "Cantidad", "Precio Unitario",
        "Importe Total", "Codigo Condicion IVA", "Importe IVA", "Concepto", "Motivo Nota",
        "Importe Otros Tributos", "Codigo Otros Tributos", "Descripcion Otros Tributos", "Base Imponible otros Tributos"
    ],
    "display": [
        "Tipo de Dato [Comprobante o Producto de un comprobante anterior]", 
        "Punto Venta C. Original [Número de punto de venta del comprobante original]", 
        "Número Comprobante Original [Número del comprobante original asociado]", 
        "Condición frente al IVA [IVA del Receptor - En excel va por código numérico]", 
        "Fecha de emisión [+-10 días del actual, si excede fallará - YYYY-MM-DD]", 
        "Fecha servicio desde [Fecha de inicio del período facturado - YYYY-MM-DD]", 
        "Fecha servicio hasta [Fecha de fin del período facturado - YYYY-MM-DD]", 
        "Fecha vencimiento pago [Fecha de vencimiento del pago de la factura - YYYY-MM-DD]", 
        "Tipo de Documento [Del que recibe la factura]", 
        "Cliente [Nombre o Razón Social]", 
        "Documento [Número sin guiones ni puntos. Ejemplo CUIT: 30423210137]", 
        "Domicilio [Lugar de residencia del que recibe la factura]",  
        "UnidadMtx [Unidades según la matriz de datos]", 
        "CodigoMtx [Código del producto (GTIN-8, GTIN-12, GTIN-13, GTIN-14)]", 
        "Descripcion [Del producto facturado]", 
        "Cantidad [Del producto facturado]", 
        "Precio Unitario [Del producto facturado]", 
        "Importe Total [Se calcula: Cantidad x Precio Unitario]", 
        "Codigo Condicion IVA [Identificador numérico del tipo de IVA]", 
        "Importe IVA [Calculado con Importe Total y Código Condición IVA]", 
        "Concepto [Código de concepto - 1, 2 o 3]", 
        "Motivo Nota [Motivo de la nota de crédito o débito, si aplica]", 
        "Importe Otros Tributos [Importe de tributos adicionales aplicados]", 
        "Codigo Otros Tributos [Código del tributo aplicado]", 
        "Descripcion Otros Tributos [Nombre del tributo aplicado]", 
        "Base Imponible otros Tributos [Importe gravado del tributo]"
    ]
}

# Definición de columnas
columns_factura_b_debito = {
    "excel": [
        "Tipo Dato", "Punto Venta C. Original", "Número Comprobante Original", "Condicion frente al IVA", "Fecha",
        "Fecha servicio desde", "Fecha servicio hasta", "Fecha vencimiento pago", "Tipo Doc", "Cliente",
        "Documento", "Domicilio", "UnidadMtx", "codigoMtx", "Descripcion", "Cantidad", "Precio Unitario",
        "Importe Total", "Codigo Condicion IVA", "Importe IVA", "Concepto", "Motivo Nota",
        "Importe Otros Tributos", "Codigo Otros Tributos", "Descripcion Otros Tributos", "Base Imponible otros Tributos"
    ],
    "display": [
        "Tipo de Dato [Comprobante o Producto de un comprobante anterior]", 
        "Punto Venta C. Original [Número de punto de venta del comprobante original]", 
        "Número Comprobante Original [Número del comprobante original asociado]", 
        "Condición frente al IVA [IVA del Receptor - En excel va por código numérico]", 
        "Fecha de emisión [+-10 días del actual, si excede fallará - YYYY-MM-DD]", 
        "Fecha servicio desde [Fecha de inicio del período facturado - YYYY-MM-DD]", 
        "Fecha servicio hasta [Fecha de fin del período facturado - YYYY-MM-DD]", 
        "Fecha vencimiento pago [Fecha de vencimiento del pago de la factura - YYYY-MM-DD]", 
        "Tipo de Documento [Del que recibe la factura]", 
        "Cliente [Nombre o Razón Social]", 
        "Documento [Número sin guiones ni puntos. Ejemplo CUIT: 30423210137]", 
        "Domicilio [Lugar de residencia del que recibe la factura]",  
        "UnidadMtx [Unidades según la matriz de datos]", 
        "CodigoMtx [Código del producto (GTIN-8, GTIN-12, GTIN-13, GTIN-14)]", 
        "Descripcion [Del producto facturado]", 
        "Cantidad [Del producto facturado]", 
        "Precio Unitario [Del producto facturado]", 
        "Importe Total [Se calcula: Cantidad x Precio Unitario]", 
        "Codigo Condicion IVA [Identificador numérico del tipo de IVA]", 
        "Importe IVA [Calculado con Importe Total y Código Condición IVA]", 
        "Concepto [Código de concepto - 1, 2 o 3]", 
        "Motivo Nota [Motivo de la nota de crédito o débito, si aplica]", 
        "Importe Otros Tributos [Importe de tributos adicionales aplicados]", 
        "Codigo Otros Tributos [Código del tributo aplicado]", 
        "Descripcion Otros Tributos [Nombre del tributo aplicado]", 
        "Base Imponible otros Tributos [Importe gravado del tributo]"
    ]
}
# Definición de columnas
columns_factura_c = {
    "excel": [
        "Tipo Dato", "Fecha", "Periodo Desde", "Periodo Hasta", "Condicion frente al IVA",
        "Tipo Doc", "Cliente", "Documento", "Domicilio", "Cant.", "Descripcion", "$ Unit.",
        "Total", "Condición Venta", "Unidad Medida", "% Bonificación", "Importe Bonificación",
        "Importe Op Ex", "Importe IVA", "Importe Tributos", "Observaciones", "Concepto",
        "Codigo Otros Tributos", "Descripcion Otros Tributos", "Base Imponible otros Tributos"
    ],
    "display": [
        "Tipo de Dato [Comprobante o Producto de un comprobante anterior]", 
        "Fecha de emisión [+-5 días del actual, si excede fallará - YYYY-MM-DD]",
        "Periodo Desde [Periodo de Facturación: YYYY-MM-DD Año-Mes-Día]", 
        "Periodo Hasta [Perdiodo de Facturación: YYYY-MM-DD Año-Mes-Día]", 
        "Condición frente al IVA [IVA del Receptor - En excel va por código numérico]",
        "Tipo de Documento [Del que recibe la factura]", 
        "Cliente [Nombre o Razón Social]", 
        "Documento [Número sin guiones ni puntos. un Ejemplo de CUIT: 30423210137]", 
        "Domicilio [Lugar de residencia del que recibe la factura]", 
        "Cantidad [Del producto facturado]", 
        "Descripcion [Del producto facturado]", 
        "Precio Unitario [Del producto facturado]",
        "Importe Total [Se calcula: Cantidad x Precio Unitario]", 
        "Condición Venta [En el PDF de la factura]", 
        "Unidad Medida [En el PDF de la factura]", 
        "Porcentaje de Bonificación [En el PDF - Ejemplo: 10%]",  
        "Importe Bonificación [Importe bonificado o descuento]",  
        "Importe Operaciones Exentas [Monto de operaciones exentas]",  
        "Importe IVA [Calculado con Importe Total y Código Condición IVA]", 
        "Importe Tributos [Importe de tributos adicionales aplicados]", 
        "Observaciones [Dejar en blanco si se desconoce]", 
        "Concepto [Código de concepto - 1, 2 o 3]",
        "Codigo Otros Tributos [Código del tributo aplicado]", 
        "Descripcion Otros Tributos [Nombre del tributo aplicado]", 
        "Base Imponible otros Tributos [Importe gravado del tributo]"
    ]
}

# Definición de columnas
columns_factura_c_credito = {
    "excel": [
        "Tipo Dato", "Número Comprobante Original", "Fecha", "Tipo Doc", "Cliente",
        "Documento", "Domicilio", "Importe Total", "Concepto", "Motivo Nota"
    ],
    "display": [
        "Tipo de Dato [Comprobante o Producto de un comprobante anterior]", 
        "Número Comprobante Original [Número del comprobante original asociado]", 
        "Fecha de emisión [+-5 días del actual, si excede fallará - YYYY-MM-DD]", 
        "Tipo de Documento [Del que recibe la factura]", 
        "Cliente [Nombre o Razón Social]", 
        "Documento [Número sin guiones ni puntos. Ejemplo CUIT: 30423210137]", 
        "Domicilio [Lugar de residencia del que recibe la factura]",  
        "Importe [Importe de la nota de credito]", 
        "Concepto [Código de concepto - 1, 2 o 3]",
        "Motivo Nota [Motivo de la nota de Credito]"
    ]
}

# Definición de columnas
columns_factura_c_debito = {
    "excel": [
        "Tipo Dato", "Número Comprobante Original", "Fecha", "Tipo Doc", "Cliente",
        "Documento", "Domicilio", "Importe Total", "Concepto", "Motivo Nota"
    ],
    "display": [
        "Tipo de Dato [Comprobante o Producto de un comprobante anterior]", 
        "Número Comprobante Original [Número del comprobante original asociado]", 
        "Fecha de emisión [+-5 días del actual, si excede fallará - YYYY-MM-DD]", 
        "Tipo de Documento [Del que recibe la factura]", 
        "Cliente [Nombre o Razón Social]", 
        "Documento [Número sin guiones ni puntos. Ejemplo CUIT: 30423210137]", 
        "Domicilio [Lugar de residencia del que recibe la factura]",  
        "Importe [Importe de la nota de debito]", 
        "Concepto [Código de concepto - 1, 2 o 3]",
        "Motivo Nota [Motivo de la nota de Debito]"
    ]
}
