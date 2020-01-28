import { version } from '../../package.json';
import { Router } from 'express';
import facets from './facets';
import WordExtractor from "word-extractor";
var docxParser = require('docx-parser');


import fs from 'fs';

export default ({ config, db }) => {
	let api = Router();
	var formMap ={
		1: {
			"nombre":"Cognoms i nom",
			"nif":"NIF",
			"fecha_nacimiento":"Data de naixement",
			"genero":"Gènere",
			"nass":"NASS",
			"discapacidad":"Discapacitat",			
		},
		2:{
			"direccion":"Adreça",
			"codigo_postal":"Codi Postal",
			"ciudad":"Població",
			"comarca":"Comarca",
			"correo":"Correu electrònic",
			"telefono":"Telèfon",
			"movil":"Telèfon mòbil",		
		},
		5:{
			"empresa_razon_social":"Raó social",
			"empresa_sector":"Sector",
			"empresa_convenio":"onveni de referència",
			"empresa_cif":"CIF",
			"empresa_ss":"a la Seguretat Social",
		},
		6:{
			"empresa_num_trabajadores":"Nre. de treballadors",
			"empresa_direccion":"Adreça del centre de treball",
			"empresa_codigo_postal":"Codi Postal",
			"empresa_poblacion":"Població",
			"empresa_comarca":"Comarca"
		}
	};
	
	function extractValueFrom(text, filterText){		
		return text.split("FORM")[0].trim().split("\u0013")[0].split("\u0007")[0].split("\r")[0];
	}
	
	api.get('/', (req, res) => {
		parseX("temp1580245941633.docx",res);
		
	});
	
	function parse(file,res){
		var extractor = new WordExtractor();
		var extracted = extractor.extract(file);
		console.log(extracted);
		extracted.then(function(doc) {
			var response={};
			var fullText = doc.pieces.map(piece=>piece.text).join("");
			//console.log(doc.pieces[5])
			Object.keys(formMap).forEach(index =>{
				var piece = doc.pieces[index].text;
				Object.keys(formMap[index]).forEach(field =>{
					var filterText = formMap[index][field];
					fullText = fullText.substring(fullText.indexOf(filterText)+filterText.length, fullText.length);
					//console.log(filterText, fullText.substring(0,15));
					response[field] = extractValueFrom(fullText, filterText);
				});
			});
			console.log("Cargado alumno ", response.nombre);
			fs.unlink(file, (err)=>{console.log(err)});
			res.json(response)
		}).catch(console.log);
	}

	api.post('/parser/doc', (req, res) => {
		var filePath="temp"+(new Date().valueOf())+".docx"
		fs.appendFile(filePath, req.files.doc.data, function() {
			parse(filePath,res);
		});
		
		
	});
	
	var formMapX ={
		8: ["nombre","Cognoms i nom"],
		10: ["nif","NIF"],
		11: ["fecha_nacimiento","Data de naixement"],
		12: ["genero","Gènere"],
		13: ["nass","NASS"],
		//"discapacidad","Discapacitat",
		15: ["direccion","Adreça"],
		16: ["codigo_postal","Codi Postal"],
		17: ["ciudad","Població"],
		18: ["comarca","Comarca"],
		20: ["correo","Correu electrònic"],
		22: ["telefono","Telèfon"],
		21: ["movil","Telèfon mòbil"],
		74: ["empresa_razon_social",""],
		75: ["empresa_sector","Sector"],
		76: ["empresa_convenio","Conveni de referència"],
		77: ["empresa_cif","CIF"],
		78: ["empresa_ss","Núm. d’inscripció a la Seguretat Social"],
		72: ["empresa_num_trabajadores","Nombre de treballadors"],
		80: ["empresa_direccion",""],
		81: ["empresa_codigo_postal","Codi Postal"],
		82: ["empresa_poblacion","Població"],
		83: ["empresa_comarca","Comarca"]
	};
	
	function parseX(file,res){
		docxParser.parseDocx(file, function(data){
			var response={};
			data.split("\n").map(console.log);
			var split = data.split("\n");
			Object.keys(formMapX).forEach(index =>{
				var field=formMapX[index][0];
				var pattern=formMapX[index][1];
				var value = split[index];
				response[field] = value.substr(pattern.length).trim();	
				if("FORMTEXT" == response[field]){
					response[field] = "";
				}					
			});
			console.log("Cargado alumno ", response.nombre);
			fs.unlink(file, (err)=>{console.log(err)});
			res.json(response)
		});
	}
	
	api.post('/parser/docx', (req, res) => {
		var filePath="temp"+(new Date().valueOf())+".docx"
		fs.appendFile(filePath, req.files.doc.data, function() {
			parseX(filePath,res);
		});
		
		
	});
	 
	return api;
}
