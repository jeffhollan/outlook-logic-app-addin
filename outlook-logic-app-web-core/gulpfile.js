'use strict';

var gulp = require('gulp');
var sass = require('gulp-sass');
var concat = require('gulp-concat');
var uglify = require('gulp-uglify');
var lib = require('bower-files')();

gulp.task('sass', ['clean'], function (callback) {

         gulp.src('**/*.scss')
          .pipe(sass().on('error', sass.logError))
          .pipe(concat('lib.css'))
          .pipe(gulp.dest('./wwwroot/dist/css'));
         callback();
    });
   
gulp.task('bower', ['sass'], function (callback) {
    gulp.src(lib.ext('js').files)
        .pipe(concat('modules.js'))
        .pipe(gulp.dest('./wwwroot/dist/js'));

    gulp.src('./assets/js/**/*.js')
        .pipe(gulp.dest('./wwwroot/dist/js/lib'));
    callback();
});

gulp.task('clean', require('del').bind(null, ['./wwwroot/dist']));

gulp.task('uglify', ['sass', 'bower'], function () {

});


gulp.task('default', function () {
    return gulp.start('bower');
});