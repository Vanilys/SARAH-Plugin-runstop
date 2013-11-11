exports.init = function (config, SARAH){
  state = 'on';
}

exports.cron = function(callback, config, task){
	    
  var exec = require('child_process').exec;
  var process = '%CD%/plugins/runstop/bin/UpTime.vbs';
  //console.log("Process CRON : " + process);
  
  var child = exec(process,
  	function (error, stdout, stderr) {
		/*infos(config);
		callback({'tts': ""});*/
		});
		
  setTimeout((function() {
	infos(config);
	}), 1000);
	
	callback({'tts': ""});
}

var dataInfos;
var infos = function (config) {
	
	var fs = require('fs');
	var file = 'plugins/runstop/uptime.json';
	dataInfos = fs.readFileSync(file,'utf8');
	dataInfos = JSON.parse(dataInfos);
	
	return dataInfos;	
};
exports.infos = infos;


var state;
var status = function(config, SARAH){
  return state;
}
exports.status = status;


exports.action = function(data, callback, config, SARAH){
   
	var exec = require('child_process').exec;
	
	need = data.need.toLowerCase();
	switch (need)
	{
	case 's_uptimesarah':
		callback({'tts': dataInfos.infos.runstop.UptimeSarah});
		var process = "default";
	break;
	case 's_uptimesystem':	
		callback({'tts': dataInfos.infos.runstop.UptimeSystem});
		var process = "default";
	break;
	
	case 's_standby':
		state = 'off';
		callback({'tts': "Mode en veille activai."});
		SARAH.remote({'context' : 'lazyRunstop.xml'});
		var process = "default";
	break;
	case 's_wakeup':
		state = 'on';
		SARAH.remote({'context' : 'default'});
		callback({'tts': "Je suis de retour !"});
		var process = "default";
	break;
	case 's_status_on':
		state = 'on';
		var process = "default";
	break;
	case 's_status_off':
		state = 'off';
		var process = "default";
	break;
	case 's_status':	
		callback({'tts': status(config, SARAH)});
		var process = "default";
	break;
	
	case 's_actions':
		var process = "SARAH_RunActions.vbs";
	break;
	
	case 's_stop':
		var process = "SARAH_Stop.vbs";
	break;
	case 's_restart':
		var process = "SARAH_Restart.vbs";
	break;
	
	case 'pc_lock':
		var process = "PC_Lock.vbs";
	break;
	case 'pc_logoff':
		var process = "PC_LogOff.vbs";
	break;
	case 'pc_sleep':
		var process = "PC_Sleep.vbs";
	break;
	case 'pc_hibernate':
		var process = "PC_Hibernate.vbs";
	break;
	case 'pc_stop':
		var process = "PC_Stop.vbs";
	break;
	case 'pc_restart':
		var process = "PC_Restart.vbs";
	break;
	case 'pc_restart_force':
		var process = "PC_RestartForce.vbs";
	break;
	case 'pc_cancelstop':
		var process = "PC_CancelStop.vbs";
	break;
	default:
		var process = "default";
	}
	
	if ((process != "default") && (process != "status")) {
		var process = '%CD%/plugins/runstop/bin/' + process;
	  	console.log("Process : " + process);
	  
	  var child = exec(process,
	  	function (error, stdout, stderr) {
			if (error !== null) console.log('exec error: ' + error);
		});
	}
				    
	switch (need)
	{
	case 's_stop':
		callback({'tts': "Au revoir."});
	break;
	case 's_restart':
		callback({'tts': "Je redaimarre."});
	break;
	
	case 'pc_lock':
		callback({'tts': "Je verrouille la session."});
	break;
	case 'pc_logoff':
		callback({'tts': "Je ferme la session."});
	break;
	case 'pc_sleep':
		callback({'tts': "Je mets l'ordinateur en veille."});
	break;
	case 'pc_hibernate':
		callback({'tts': "Je mets l'ordinateur en hibernation."});
	break;
	case 'pc_stop':
		callback({'tts': "J'arraite l'ordinateur."});
	break;
	case 'pc_restart':
		callback({'tts': "Je redaimarre l'ordinateur."});
	break;
	case 'pc_restart_force':
		callback({'tts': "Je force l'ordinateur a redaimarrai."});
	break;
	case 'pc_cancelstop':
		callback({'tts': "J'annulle l'arrai de l'ordinateur."});
	break;
	default:
		callback({'tts': ""});
	}
	
}
