# caser-servicios
Se trata de un programa que funciona con integración de Gmail y Googlesheet. Mediante AppScript y AppSheet se manejan todas las necesidades de la empresa para gestionar los servicios, las citas, los presupuestos asociados y un sistema de facturación para contabilidad. Además tiene una agenda y un sistema de registro de asistencia de operarios.

La empresa para la que fue diseñado este programa en concreto es una empresa que colabora con CASER para realizar los servicios del hogar. La compañía le envía un correo electrónico a la empresa en el que se incluye un AVISO adjunto en un PDF con todos los datos del siniestro: número de servicio, información del tomador, dirección del siniestro, descripción del hecho, teléfonos de contacto, incluso también puede venir información de cita.

El programa se encarga de guardar el correo en su etiqueta pertinente según la información del correo recibida y una vez etiquetado procesa sus datos para incluirlos en la hoja de sheet correspondiente:
  - En caso de ser un servicio nuevo, en la hoja SERVICIOS: inserta la fecha de alta (hoy), la categoría de la póliza, el código ref de servicio, observaciones (puede ser urgencia), fecha de cita, hora de cita, estado "CITAR" (o "CITADO") si que tiene fijada cita por defecto, además de dirección y localidad, descripción del servicio y teléfonos de contacto. También recoge dni y nombre del tomador de la póliza asociada al siniestro.
  - En caso de ser un servicio reaperturado, cambia el estado del servicio introducido, que debía ser "CERRADO", por "CITAR", actualiza la fecha de cita para que sea un siniestro pendiente del día siguiente y cambia la descripción si es que hay un texto explícito indicativo para nueva descripción.
  - En caso de recibir una incidencia, etiqueta el correo y lo deja no leído para gestionar manualmente el aviso. A modo de notificación se queda así para mayor visibilidad. Configurable automáticamente.
  - En caso de recibir un seguimiento (mensaje informativo respecto de un servicio), etiqueta el correo y lo deja leído, ya que este pasa a la hoja sheet de SEGUIMIENTOS, lo vincula al servicio correspondiente para que sea accesible desde AppSheet.
  - En caso de recibir una devolución de factura, etiqueta el correo y lo deja no leído para gestionar manualmente el aviso. A modo de notificación se queda así para mayor visibilidad. Configurable automáticamente.
  - En caso de recibir una notificación de cita posterior a la recepción del aviso, se encarga de actualizar dicho registro de la hoja SERVICIOS, marcando "CITADO" en estado y asignando fecha y hora de la cita.



(EN PROCESO)
Se encuentra en proceso la optimización de gestión de presupuestos, que actualmente requiere creación desde googlesheet en una página que tiene configurados botones para añadir tarifas asociadas.
Está pendiente de implementarse que desde AppSheet se pueda crear un presupuesto y este pueda ser cargado directamente en una plantilla de doc sin necesidad del paso intermedio de aparecer en una plantilla de googlesheet como está actualmente configurado.

  - - 
