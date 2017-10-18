edge-vb
=======

This is a VB compiler for edge.js.

Please fork on GitHub if you think you can help improve this project.

Example "Hello World" code:

	var WriteCrapVB = edge.func('vb', function () {/*
        Async Function(Input As Object) As Task(Of Object)
            Return Await Task.Run(Function()
				Return "NodeJS Welcomes: " & Input.ToString()
            End Function)
        End Function
    */});
    WriteCrapVB('VB', function (error, result) {
        if (error) throw error;
        console.log(result); // Returns "NodeJS Welcomes: VB"
    });

See [edge.js overview](http://tjanczuk.github.com/edge) and [edge.js on GitHub](https://github.com/tjanczuk/edge) for more information. 
