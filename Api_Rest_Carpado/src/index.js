const mqtt = require('mqtt');
const sql = require('mssql');

// Base de Datos SQL SERVER
const config = {
    user: 'sisma_app',
    password: 'V$123bcd',
    database: 'Tracer',
    server: 'LocalHost',
    pool: {
        max: 10,
        min: 0,
        idleTimeoutMillis: 1000
    },
    options: {
        encrypt: true, // for azure
        trustServerCertificate: true // change to true for local dev / self-signed certs
    }
};

//CREDENCIALES MQTT
var options = {
  port: 1883,
  host: '18.229.172.128',
  clientId: 'CIFPL_CARPADO_' + Math.round(Math.random() * (0- 10000) * -1) ,
  username: 'web_client',
  password: 'Edgar6305',
  keepalive: 60,
  reconnectPeriod: 1000,
  protocolId: 'MQIsdp',
  protocolVersion: 3,
  clean: true,
  encoding: 'utf8'
};

var client = mqtt.connect("mqtt://18.229.172.128", options);

//SE REALIZA LA CONEXION
client.on('connect', function () {
  console.log("Conexión  MQTT Exitosa!");
  client.subscribe('cifpl/app/carpado/datos', function (err) {
    console.log("Subscripción exitosa!")
  });
})

//CUANDO SE RECIBE MENSAJE
client.on('message', function (topic, message) {
  var message_splitted = message.toString().split("/")
  var tiquete = message_splitted[0]
  var placa = message_splitted[1]
  //console.log("Mensaje recibido desde -> " + topic + " Mensaje -> " + message.toString());
  //console.log("Tiquete ==> " + tiquete + " Placa " + placa)
  const r = consulta(tiquete, placa)
});


// Consulta a SQL SERVER 2019
async function  consulta(tiquete, placascarpado){
    try {
        await sql.connect(config)
        var xSql =`Select Placas FROM Bascula WHERE IdTiquete=${tiquete} AND Estado='AC'`
        var result = await sql.query (xSql)
        if (result.rowsAffected[0] == 0){
            client.publish('cifpl/app/carpado/respuesta',"Tiquete NO Localizado, Verifique  No=> " + tiquete);
        }else{
            //console.log(result.recordset[0])
            var placas=result.recordset[0].Placas
            if(placascarpado == placas) {
                xSql =`Update Bascula Set Carpado=1 WHERE IdTiquete=${tiquete}`
                var resultUP = await sql.query (xSql)
                client.publish('cifpl/app/carpado/respuesta',"Carpado Actualizado Tiquete No=> " + tiquete);
                xSql =`INSERT INTO LogCarpado VALUES(${tiquete},'${placas}', Getdate())`
                //console.log(xSql)
                resultUP = await sql.query (xSql)
                console.log("Registro Actualizado...")
            }else{
                client.publish('cifpl/app/carpado/respuesta',"La PLACA NO concuerda con el tiquete, Verifique");
            }
        }    
        return        
    } catch (err) {
        client.publish('cifpl/app/carpado/respuesta',err.message);
    }    
}

setInterval(function () {
  //client.publish('cifpl/app/carpado/server',"Emitiendo");    
  }, 60000);
 