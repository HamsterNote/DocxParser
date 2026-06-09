/** @type {import('jest').Config} */
module.exports = {
  preset: 'ts-jest',
  testEnvironment: 'node',
  roots: ['<rootDir>/src'],
  // 禁用 Jest 30 的全局变量清理机制（globalsCleanup: 'soft' 存在已知 bug）
  // 参见：https://github.com/jestjs/jest/issues/16044
  testEnvironmentOptions: {
    globalsCleanup: 'off'
  },
  // 将 .ts 文件作为 ESM 处理
  extensionsToTreatAsEsm: ['.ts'],
  moduleNameMapper: {
    '^(\\.{1,2}/.*)\\.js$': '$1',
    '^@/(.*)$': '<rootDir>/src/$1',
    '^src/(.*)$': '<rootDir>/node_modules/@hamster-note/types/src/$1',
    // 映射 ESM 模块到其编译后的入口文件
    '^@hamster-note/document-parser$':
      '<rootDir>/node_modules/@hamster-note/document-parser/dist/index.js',
    '^@hamster-note/types$':
      '<rootDir>/node_modules/@hamster-note/types/dist/index.js'
  },
  transform: {
    // 转换 TypeScript 文件
    '^.+\\.(ts|tsx)$': [
      'ts-jest',
      {
        tsconfig: {
          ...require('./tsconfig.jest.json').compilerOptions,
          allowJs: true,
          checkJs: false
        },
        isolatedModules: false,
        useESM: true
      }
    ],
    // 转换 node_modules 中 @hamster-note 包的 ESM JS 文件
    'node_modules/@hamster-note/.+\\.js$': [
      'ts-jest',
      {
        tsconfig: 'tsconfig.jest.json',
        isolatedModules: true,
        useESM: true
      }
    ]
  },
  collectCoverage: true,
  coverageDirectory: 'coverage',
  coverageReporters: ['text', 'lcov', 'html'],
  // 不忽略 @hamster-note 包，允许 Jest 转换它们
  transformIgnorePatterns: ['/node_modules/(?!(@hamster-note)/)']
}
