var gulp = require('gulp');
var sourcemaps = require('gulp-sourcemaps');
var ts = require('gulp-typescript');
var concat = require('gulp-concat');
var tsproject = require( 'tsproject' );

var compilerOptions = require("./tsconfig.json").compilerOptions;

var jsDir = __dirname + "/dist";
var tsJsFileName = "shgeneric.js";

gulp.task( 'ts-common', function() {
    return tsproject.src('./tsconfig.json',
        {
            logLevel: 1,
            compilerOptions: {
                listFiles: true
            }
        })
        .pipe( gulp.dest( './dist' ) );
});


gulp.task('default', ['ts-common']);
