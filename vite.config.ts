import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';
import path from 'path';
import { viteStaticCopy } from 'vite-plugin-static-copy';

import fs from 'fs';
import os from 'os';

const certPath = path.join(os.homedir(), '.office-addin-dev-certs', 'localhost.crt');
const keyPath = path.join(os.homedir(), '.office-addin-dev-certs', 'localhost.key');
const hasCerts = fs.existsSync(certPath) && fs.existsSync(keyPath);

export default defineConfig({
  plugins: [
    react(),
    viteStaticCopy({
      targets: [
        {
          src: 'node_modules/katex/dist/fonts/*',
          dest: 'assets/katex-fonts'
        },
        {
          src: 'node_modules/katex/dist/fonts/*',
          dest: 'assets/fonts'
        },
        {
          src: 'node_modules/katex/dist/fonts/*',
          dest: 'fonts'
        },
        {
          src: 'node_modules/katex/dist/katex.min.css',
          dest: 'assets'
        },
        {
          src: 'node_modules/mathlive/dist/fonts/*',
          dest: 'assets/mathlive-fonts'
        }
      ]
    })
  ],
  base: './',
  resolve: {
    alias: {
      '@': path.resolve(__dirname, './src')
    }
  },
  server: {
    port: 3000,
    https: hasCerts ? {
      key: fs.readFileSync(keyPath),
      cert: fs.readFileSync(keyPath.replace('.key', '.crt'))
    } : undefined,
    headers: {
      'Access-Control-Allow-Origin': '*'
    }
  },
  preview: {
    port: 3000,
    https: hasCerts ? {
      key: fs.readFileSync(keyPath),
      cert: fs.readFileSync(keyPath.replace('.key', '.crt'))
    } : undefined,
    headers: {
      'Access-Control-Allow-Origin': '*'
    }
  },
  build: {
    outDir: 'dist',
    emptyOutDir: true,
    sourcemap: true,
    rollupOptions: {
      input: {
        taskpane: path.resolve(__dirname, 'taskpane.html'),
        functions: path.resolve(__dirname, 'src/customfunctions/index.ts')
      },
      output: {
        entryFileNames: 'assets/[name].js',
        chunkFileNames: 'assets/[name]-[hash].js',
        assetFileNames: 'assets/[name].[ext]',
        manualChunks: {
          'react-vendor': ['react', 'react-dom'],
          'fluentui': ['@fluentui/react-components', '@fluentui/react-icons']
        }
      }
    }
  }
});
