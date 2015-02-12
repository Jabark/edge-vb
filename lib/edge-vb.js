var path = require('path');

exports.getCompiler = function () {
	return process.env.EDGE_VB_NATIVE || path.join(__dirname, 'edge-vb.dll');
};
