var gulp = require('gulp');
var sourcemaps = require('gulp-sourcemaps');
var ts = require('gulp-typescript');
var concat = require('gulp-concat');

var compilerOptions = require("./tsconfig.json").compilerOptions;

var jsDir = __dirname + "/dist";
var tsJsFileName = "shgeneric.js";

gulp.task('typescript-compile', function(){
    var tsResult = gulp.src('TS/*.ts')
                       .pipe(sourcemaps.init()) // This means sourcemaps will be generated 
                       .pipe(ts(compilerOptions));
    
    return tsResult.js
                .pipe(concat(tsJsFileName))
                .pipe(sourcemaps.write())
                .pipe(gulp.dest(jsDir));
});

gulp.task('default', ['typescript-compile']);
