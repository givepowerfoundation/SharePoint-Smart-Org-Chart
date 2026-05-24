'use strict';

const build = require('@microsoft/sp-build-web');

build.addSuppression(
  `Warning - [sass] The local CSS class 'ms-Grid' is not camelCase and will not be type-safe.`
);

const getTasks = build.task('configure-webpack-logging', {
  execute: () => Promise.resolve()
});

build.configureWebpack.mergeConfig({
  additionalConfiguration: (generatedConfiguration) => {
    return generatedConfiguration;
  }
});

const gulp = require('gulp');
build.initialize(gulp);

gulp.task('serve', gulp.series('serve-deprecated'));
