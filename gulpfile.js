var gulp = require("gulp");
var concat = require("gulp-concat");
var uglify = require("gulp-uglify");
var wrap = require("gulp-wrap");

gulp.task("concat", function()
{
    return gulp.src("src/**/*.js")
        .pipe(concat("oneNoteToMarkdown.js"))
        .pipe(wrap({
            src: "build-template.js",
            parse: false
        }))
        .pipe(gulp.dest("dist"));
});

gulp.task("minify", function()
{
    return gulp.src("src/**/*.js")
        .pipe(concat("oneNoteToMarkdown.min.js"))
        .pipe(uglify())
        .pipe(wrap({
            src: "build-template.js",
            parse: false
        }))
        .pipe(gulp.dest("dist"));
});

gulp.task("watch", function()
{
    gulp.watch("src/**/*.js", ["concat", "minify"]);
});

gulp.task("default", ["watch"]);