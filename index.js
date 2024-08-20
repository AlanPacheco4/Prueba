const { log } = require('console');
const { Octokit, App } = require ('octokit');
const XlsxPopulate = require ('xlsx-populate');
const { name } = require('xlsx-populate/lib/RichText');

const octokit = new Octokit({ 
  baseUrl: "https://github.banco.gfi.mx/api/v3",
  auth: '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'
});

//Funcion Fecha para el nombre del archivo
function obtfecha (){
  const date = new Date();
  const [day, month, year] =[
    date.getDate(),
    date.getMonth(),
    date.getFullYear()
  ];
  const fecha = (day+"-"+(month+1)+"-"+year);

  return fecha
}

obtfecha();

//Listar repos
async function listrepos(){

    const orgn = 'GrupoFinancieroInbursa';
    const response = await octokit.paginate(octokit.rest.dependabot.listAlertsForOrg,{
          org: orgn,
          direction: "asc",
          //per_page:100, //no es necesario paginar, la funcion asincorna paginate devuelve toda la info
    });

    const repos =[];
    for(var i in response){
      repos [i] = response[i].repository.name;
        //console.log(response[i]);
    }
    let filtro = repos.filter((item,index)=>{
      return repos.indexOf(item) === index;
    })

    

    for (let i in filtro){
      console.log(filtro[i]);

      const repos = filtro[i];
      //Consulta de la Api
      const respons = await octokit.paginate(octokit.rest.dependabot.listAlertsForRepo, {
        owner: orgn,
        repo: repos,
        direction: "asc",
      });

      let iter = respons;
    
      XlsxPopulate.fromFileAsync('./plantillav2.xlsx')
      .then(workbook => {
    
        workbook.sheet("Vulnerabilidades").cell("A2").value("DEPENDABOT ALERTS: " + repos);
        workbook.sheet("Vulnerabilidades").cell("K1").value("Alertas: " + iter.length);
        workbook.sheet("Vulnerabilidades").cell("L1").value("Fecha reporte: " + obtfecha());
        let c = 4;
        
        //Iteracion de la consulta e inserccion de valores en Tabla
        for (var i in iter){
          let fecha = iter[i].created_at;
          let date = new Date(fecha).toLocaleString("es-MX");
    
          let fixeo;
          let fech_fix = iter[i].fixed_at;
          if(fech_fix == null){
            fixeo ="Sin fecha";
          }
          else{
            fixeo = new Date(fech_fix).toLocaleString("es-MX");
          }       
          //console.log(fecha.slice(0, 10));
          let letra = "I";
          //console.log(iter[i].security_advisory.cvss.score);
          workbook.sheet("Vulnerabilidades").cell("A"+ (c)).value(iter[i].number);
          workbook.sheet("Vulnerabilidades").cell("B"+ (c)).value(iter[i].security_advisory.summary);
          //workbook.sheet("Vulnerabilidades").cell("B"+ (c)).value("Ecosistema: "+iter[i].dependency.package.ecosystem+"        Nombre: "+iter[i].dependency.package.name);
          workbook.sheet("Vulnerabilidades").cell("C"+ (c)).value(iter[i].state);
          workbook.sheet("Vulnerabilidades").cell("D"+ (c)).value(iter[i].security_vulnerability.severity);
          workbook.sheet("Vulnerabilidades").cell("E"+ (c)).value(iter[i].security_advisory.cvss.score);
          workbook.sheet("Vulnerabilidades").cell("F"+ (c)).value(iter[i].security_advisory.description);
          workbook.sheet("Vulnerabilidades").cell("G"+ (c)).value("Versiones afectadas: "+iter[i].security_vulnerability.vulnerable_version_range   +"           Version Parchada: "+ iter[i].security_vulnerability.first_patched_version.identifier);
          workbook.sheet("Vulnerabilidades").cell("H"+ (c)).value("GHSA: "+iter[i].security_advisory.ghsa_id+"    |CVE:"+iter[i].security_advisory.cve_id+ "    |"+iter[i].security_advisory.cvss.vector_string+"     |Score: "+ iter[i].security_advisory.cvss.score);
          workbook.sheet("Vulnerabilidades").cell("I"+ (c)).value(date);
          workbook.sheet("Vulnerabilidades").cell("K" + (c)).value(fixeo);
          c = c+1;
    
          const pib = (iter[i].security_vulnerability.severity);
          if (pib=="low"){
            workbook.sheet("Vulnerabilidades").cell("D"+ (c-1)).style("fill", "88DA26");
          }
          if (pib=="medium"){
            workbook.sheet("Vulnerabilidades").cell("D"+ (c-1)).style("fill", "ffff00");
          }
          else if (pib=="high"){
            workbook.sheet("Vulnerabilidades").cell("D"+ (c-1)).style("fill", "EE6902");
          }
          else if (pib=="critical"){
            workbook.sheet("Vulnerabilidades").cell("D"+ (c-1)).style("fill", "F70606");
          }
    
          const stat = (iter[i].state);
          if (stat=="open"){
            workbook.sheet("Vulnerabilidades").cell("C"+ (c-1)).style("fontColor", "722F37");
          }
          if (stat=="fixed"){
            workbook.sheet("Vulnerabilidades").cell("C"+ (c-1)).style("fontColor", "116C09");
          }
        }
        console.log(repos, "Creado con exito");
        return workbook.toFileAsync("../../Reportes Dependabot/Por Repocitorio/Reporte_"+repos+"_"+obtfecha()+".xlsx");
      });

    }
    console.log("tarea finalizada");
      
}
//listrepos();




async function showrepos(){
  const orgn = 'GrupoFinancieroInbursa';
  const response = await octokit.request(
    'GET /orgs/{org}/dependabot/alerts', 
    {
      org: orgn,
      headers: {
        "x-github-api-version": "2022-11-28",
      },
    });


    
    const iter = response.data;
    //const dan = JSON.stringify(iter);


    /*let repocit = ["robert", "el Bordddddddddddddddddddd", "Brows", "robert", "alana","alan","alan"];
    let res = [...new Set(repocit) ];*/
    //console.log (res);
    //console.log(res);

    //let arr = [];
    //let itera = JSON.parse(iter);
    //console.log(itera)

      for(var i = 0; i < itera.length; i++){
        let repocit2= iter[i].repository.name;
        //console.log(i);
         //arr.push(iter[i].repository.name);

        //const result = Object.assign({}, ...repocit);
        //console.log(arr);
      }
      


}
//showrepos();


//Listar alertas por Repo con excel
async function showbyrep(){
  //Varible para repo
  const repos = 'SIIBAN-banco';

  //Consulta de la Api
  const response = await octokit.request(
    'GET /repos/{owner}/{repo}/dependabot/alerts', {
      owner: 'GrupoFinancieroInbursa',
      repo: repos,
    });
  
  //Se escoge solo los datos de la consulta
  const iter = response.data;
  //console.log (iter);
  //let dato = (iter[1].created_at).slice(0, 10);
  let dato = (iter[1].created_at);
  //let date = new Date(dato+ "GMT-6").toUTCString();
  let date = new Date(dato).toLocaleString("es-MX");

  console.log(date);
  //console.log(dato);
  //console.log(new Date());

    for (var i = (iter.length - 1); i >= 0 ; i--){    
      /*console.log(iter[i].number);
      console.log("Ecosistema: "+iter[i].dependency.package.ecosystem+"        Nombre: "+iter[i].dependency.package.name);
      console.log(iter[i].state);
      console.log(iter[i].security_vulnerability.severity);
      //console.log(iter[i].security_advisory.description);
      console.log(iter[i].security_vulnerability.vulnerable_version_range);
      console.log("GHSA: "+iter[i].security_advisory.ghsa_id+"    |CVE:"+iter[i].security_advisory.cve_id+ "    |"+iter[i].security_advisory.cvss.vector_string+"     |Score: "+ iter[i].security_advisory.cvss.score);*/
    }

  }

//showbyrep();


///////////////////////////Listar alertas de la organizacion////////////////////////////////////////////
async function showbyorg() {
  const orgn = 'GrupoFinancieroInbursa';
  const response = await octokit.paginate(octokit.rest.dependabot.listAlertsForOrg,{
    org: orgn,
    direction: "asc",  
    /*'GET /orgs/{org}/dependabot/alerts', 
    {
      org: orgn,
      per_page: 90,
      direction: "asc",
      headers: {
        "x-github-api-version": "2022-11-28",
      },*/
    });

    const iter = response;
    //console.log(iter);
  //Uso de modulo XlsPopulate
  XlsxPopulate.fromFileAsync('./plantilla_org.xlsm')
  .then(workbook => {

    //Encabezados de la tabla
    workbook.sheet("Vulnerabilidades").cell("D1").value("Alertas Dependabot por organizacion "+orgn+ ": "+iter.length);
    workbook.sheet("Vulnerabilidades").cell("F1").value("Fecha reporte: " + obtfecha());    
    //console.log(response.data);
    //Iteracion de la consulta e inserccion de valores en Tabla
    let c = 3;
    console.log(iter[1]);
    /*for (let i in iter){
      //console.log(i);
      //console.log(iter[i].repository.name);
      
      let fecha = iter[i].created_at;
      let date = new Date(fecha).toLocaleString("es-MX");
      //console.log(iter.length);
      workbook.sheet("Vulnerabilidades").cell("A"+ (c)).value(iter[i].number);
      workbook.sheet("Vulnerabilidades").cell("B"+ (c)).value(iter[i].repository.name);      
      workbook.sheet("Vulnerabilidades").cell("C"+ (c)).value(iter[i].dependency.package.ecosystem);
      workbook.sheet("Vulnerabilidades").cell("D"+ (c)).value(iter[i].state);
      workbook.sheet("Vulnerabilidades").cell("E"+ (c)).value(iter[i].security_vulnerability.severity);
      workbook.sheet("Vulnerabilidades").cell("F"+ (c)).value(date);
      c = c+1;
    }*/
    //return workbook.toFileAsync("../../Reportes Dependabot/Por Organizacion/Reporte_"+orgn+"_"+obtfecha()+".xlsm");
    
  });
}

//showbyorg();



////////////////////////////////////////agrega score
//Listar alertas por Repositorio con excel
async function showbyrep2(){
  //Varible para repo
  const ow = 'GrupoFinancieroInbursa';
  const repos = 'BLP-site-bep-js-webapp';
  //Consulta de la Api
  const response = await octokit.paginate(octokit.rest.dependabot.listAlertsForRepo, {
    owner: ow,
    repo: repos,
    direction: "asc",
  });
  
  //Se escoge solo los datos de la consulta
  const iter = response;

  //Uso de modulo XlsPopulate
  XlsxPopulate.fromFileAsync('./plantillav2.xlsx')
  .then(workbook => {


    workbook.sheet("Vulnerabilidades").cell("A2").value("DEPENDABOT ALERTS: " + repos);
    workbook.sheet("Vulnerabilidades").cell("K1").value("Alertas: " + iter.length);
    workbook.sheet("Vulnerabilidades").cell("L1").value("Fecha reporte: " + obtfecha());
    let c = 4;
    
    //Iteracion de la consulta e inserccion de valores en Tabla
    for (var i in iter){
      console.log(iter[i]);
      /*let fecha = iter[i].created_at;
      let date = new Date(fecha).toLocaleString("es-MX");

      let fixeo;
      let fech_fix = iter[i].fixed_at;
      if(fech_fix == null){
        fixeo ="Sin fecha";
      }
      else{
        fixeo = new Date(fech_fix).toLocaleString("es-MX");
      }       
      //console.log(fecha.slice(0, 10));
      let letra = "I";
      //console.log(iter[i].security_advisory.cvss.score);
      workbook.sheet("Vulnerabilidades").cell("A"+ (c)).value(iter[i].number);
      workbook.sheet("Vulnerabilidades").cell("B"+ (c)).value(iter[i].security_advisory.summary);
      //workbook.sheet("Vulnerabilidades").cell("B"+ (c)).value("Ecosistema: "+iter[i].dependency.package.ecosystem+"        Nombre: "+iter[i].dependency.package.name);
      workbook.sheet("Vulnerabilidades").cell("C"+ (c)).value(iter[i].state);
      workbook.sheet("Vulnerabilidades").cell("D"+ (c)).value(iter[i].security_vulnerability.severity);
      workbook.sheet("Vulnerabilidades").cell("E"+ (c)).value(iter[i].security_advisory.cvss.score);
      workbook.sheet("Vulnerabilidades").cell("F"+ (c)).value(iter[i].security_advisory.description);
      workbook.sheet("Vulnerabilidades").cell("G"+ (c)).value("Versiones afectadas: "+iter[i].security_vulnerability.vulnerable_version_range   +"           Version Parchada: "+ iter[i].security_vulnerability.first_patched_version.identifier);
      workbook.sheet("Vulnerabilidades").cell("H"+ (c)).value("GHSA: "+iter[i].security_advisory.ghsa_id+"    |CVE:"+iter[i].security_advisory.cve_id+ "    |"+iter[i].security_advisory.cvss.vector_string+"     |Score: "+ iter[i].security_advisory.cvss.score);
      workbook.sheet("Vulnerabilidades").cell("I"+ (c)).value(date);
      workbook.sheet("Vulnerabilidades").cell("K" + (c)).value(fixeo);
      c = c+1;

      const pib = (iter[i].security_vulnerability.severity);
      if (pib=="low"){
        workbook.sheet("Vulnerabilidades").cell("D"+ (c-1)).style("fill", "88DA26");
      }
      if (pib=="medium"){
        workbook.sheet("Vulnerabilidades").cell("D"+ (c-1)).style("fill", "ffff00");
      }
      else if (pib=="high"){
        workbook.sheet("Vulnerabilidades").cell("D"+ (c-1)).style("fill", "EE6902");
      }
      else if (pib=="critical"){
        workbook.sheet("Vulnerabilidades").cell("D"+ (c-1)).style("fill", "F70606");
      }

      const stat = (iter[i].state);
      if (stat=="open"){
        workbook.sheet("Vulnerabilidades").cell("C"+ (c-1)).style("fontColor", "722F37");
      }
      if (stat=="fixed"){
        workbook.sheet("Vulnerabilidades").cell("C"+ (c-1)).style("fontColor", "116C09");
      }*/
    }
    console.log(repos, "Creado con exito");
    //return workbook.toFileAsync("../../Reportes Dependabot/Por Repocitorio/Reporte_"+repos+"_"+obtfecha()+".xlsx");
    
});
}

showbyrep2();

