exports.action = function(data, callback, config, SARAH){
  
var exec = require('child_process').exec;
  
	if (data.need == "stop") {
		var process = "Stop_SARAH.vbs";
	}
	if (data.need == "restart") {
		var process = "Restart_SARAH.vbs";
	}
	
	var process = '%CD%/plugins/runstop/bin/' + process;
  ;console.log(process);
  
  var child = exec(process,
  	function (error, stdout, stderr) {
		});
	if (data.need == "stop") {
		callback({'tts': "Au revoir."});
	}
	if (data.need == "restart") {
		callback({'tts': "Je redaimarre."});
	}
		    
}
