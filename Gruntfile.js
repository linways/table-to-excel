module.exports = function(grunt) {
  grunt.initConfig({
    uglify: {
      options: {
        mangle: false
      },
      my_target: {
        files: {
          "dist/xlsx_html_utils.min.js": ["src/xlsx_html_utils.js"]
        }
      }
    },
    concat: {
      dist: {
        src: [
          "lib/xlsx.core.min.js",
          "lib/FileSaver.min.js",
          "dist/xlsx_html_utils.min.js"
        ],
        dest: "dist/xlsx_html.full.min.js"
      }
    },
    watch: {
      files: ["src/**/*.js"],
      tasks: ["uglify", "concat"]
    }
  });

  grunt.loadNpmTasks("grunt-contrib-uglify");
  grunt.loadNpmTasks("grunt-contrib-watch");
  grunt.loadNpmTasks("grunt-contrib-concat");

  grunt.registerTask("default", ["uglify", "concat"]);
};
