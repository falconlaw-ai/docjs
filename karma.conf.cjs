module.exports = (config) => {
  config.set({
    basePath: '',
    frameworks: ['jasmine'],
    plugins: [
      require('karma-jasmine'),
      require('karma-chrome-launcher'),
      require('karma-firefox-launcher')
    ],
    files: [
      'node_modules/jszip/dist/jszip.js',
      'node_modules/diff/dist/diff.js',
      'dist/docx-preview.js',
      'tests/**/*spec.js',
      { pattern: 'tests/**/*.docx', included: false },
      { pattern: 'tests/**/*.html', included: false }
    ],
    reporters: ['progress'],
    port: 9876,
    colors: true,
    logLevel: config.LOG_INFO,
    autoWatch: true,
    browsers: ['Chrome'],
    singleRun: false,
    concurrency: Infinity,
    crossOriginAttribute: false
  })
}
