// Generated on 2015-01-12 using generator-angular 0.10.0
'use strict';

// # Globbing
// for performance reasons we're only matching one level down:
// 'test/spec/{,*/}*.js'
// use this if you want to recursively match all subfolders:
// 'test/spec/**/*.js'

module.exports = function(grunt) {

  // Load grunt tasks automatically
  require('load-grunt-tasks')(grunt);

  // Define the configuration for all the tasks
  grunt.initConfig({
    watch: {
      server: {
        files: ['server.js'],
        tasks: ['nodemon:dev']
      }
    },
    nodemon: {
      dev: {
        script: 'server.js'
      }
    }
  });

  grunt.registerTask('default', [
    'nodemon:dev', 'watch'
  ]);
};
