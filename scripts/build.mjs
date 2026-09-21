import { build } from 'esbuild';
import { mkdir, cp, rm } from 'node:fs/promises';
await rm('dist',{recursive:true,force:true}); await mkdir('dist',{recursive:true});
const common={bundle:true,format:'esm',target:'chrome110',sourcemap:false,logLevel:'info'};
await Promise.all([
 build({...common,entryPoints:['src/content/content.ts'],outfile:'dist/content.js'}),
 build({...common,entryPoints:['src/content/background.ts'],outfile:'dist/background.js'}),
 build({...common,entryPoints:['src/popup/popup.ts'],outfile:'dist/popup.js'}),
 cp('src/popup/popup.css','dist/popup.css')
]);
