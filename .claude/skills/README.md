# Skills de UI e desenvolvimento web

55 skills de terceiros para o Claude Code, focadas em design de interface, animação, acessibilidade e qualidade web. Cada pasta tem um `SKILL.md` e a licença original (MIT ou Apache-2.0).

## Como funcionam

- **Automático:** o Claude lê a descrição de cada skill e carrega a skill inteira quando o pedido combina com ela. Por exemplo, "deixa essa animação mais suave" tende a acionar as skills de animação.
- **Pelo nome:** digite `/nome-da-skill` (ex.: `/break-ui`, `/impeccable audit`) ou escreva "use a skill X". Algumas skills só rodam assim.
- **Onde valem:** estando em `.claude/skills/`, funcionam quando o Claude Code roda **neste repositório**. Para ter em **todos os projetos**, copie as pastas para `~/.claude/skills/`:
  ```bash
  cp -r .claude/skills/* ~/.claude/skills/
  ```
- **Custo:** o nome e a descrição de cada skill ficam no contexto em toda conversa, cerca de 5 mil tokens para as 55. O texto completo só carrega quando a skill é usada.

## Tabela

### Direção estética (escolha uma por tarefa)
| Skill | O que é | Como acionar |
|---|---|---|
| `impeccable` | Pacote completo de design com 24 comandos: planejar, criar, criticar, auditar e polir | `/impeccable init` (uma vez), depois `/impeccable shape`, `craft`, `critique`, `audit`, `polish`, `bolder`, `quieter`, `typeset`, `layout`, `animate`, `harden`… |
| `frontend-design` | Skill oficial da Anthropic: escolhe uma direção estética antes de escrever código | Automática ao pedir uma página ou componente novo |
| `design-taste-frontend` | Taste Skill v2: sorteia uma direção de arte, busca imagens com licença aberta e confere em navegador real. **Leva 30–45 min e muitos tokens; precisa de Node 22+ e Chrome** | `/design-taste-frontend` para landing pages caprichadas |
| `design-taste-frontend-v1` | Taste Skill v1, mais leve e direta | Só pelo nome |
| `high-end-visual-design` | Visual calmo e refinado, com bastante respiro e animação com mola | Só pelo nome |
| `minimalist-ui` | Editorial monocromático quente, grids bento | Só pelo nome |
| `industrial-brutalist-ui` | Brutalista: tipografia suíça e grid rígido | Só pelo nome |
| `ui-ux-pro-max` | Banco de dados com 79 estilos, 192 paletas e 74 pares de fontes; gera um design system por tipo de produto. **Precisa de Python 3** | Automática, ou "use ui-ux-pro-max para um app de clínica" |
| `interface-design` | Mantém a consistência de UI de produto (dashboards, admin) salvando decisões em `.interface-design/system.md` | `/interface-design` |
| `theme-factory` | 10 temas prontos (cores e fontes) aplicáveis a páginas e slides | "aplica um tema do theme-factory" |
| `brand-guidelines` | Exemplo da Anthropic de identidade visual em forma de skill. Serve de **modelo** para criar a skill da sua marca | Para referência |

### Animação e movimento (Emil Kowalski)
| Skill | O que é | Como acionar |
|---|---|---|
| `emil-design-eng` | A principal: filosofia de acabamento, componentes e decisões de animação | Automática em trabalho de UI |
| `find-animation-opportunities` | Aponta onde vale animar e onde não. Só lê, não altera nada | `/find-animation-opportunities` |
| `animate` | Cria uma animação do zero: curva, duração, propriedades | `/animate` ou "anima esse modal" |
| `review-animations` | Revisão rigorosa de animações existentes | `/review-animations` |
| `improve-animations` | Audita o projeto todo e gera planos de correção | `/improve-animations` |
| `animation-performance` | Corrige animações que travam ou perdem frames | Automática ao reclamar de travamento |
| `animation-vocabulary` | Traduz "aquele efeito que quica" para o termo técnico certo | "como se chama o efeito…" |
| `apple-design` | Princípios de interface e movimento da Apple adaptados para a web | "estilo Apple", gestos, molas |
| `mobile-native` | Faz o site parecer app no celular: 100vh, hover preso, zoom em input, safe areas | `/mobile-native` |
| `prototype` | Gera várias versões de um componente e mostra um seletor para trocar entre elas | `/prototype um card de preço` |
| `break-ui` | Testa a UI com dados extremos: nomes enormes, listas vazias, números gigantes | `/break-ui` |
| `pick-ui-library` | Recomenda a biblioteca certa (gráficos, drag and drop, OTP…) | "qual biblioteca usar para…" |
| `ask-sonner` | Guia do Sonner (toasts em React) | Automática ao usar Sonner |

### Acabamento e revisão de interface (Jakub Krehel)
| Skill | O que é | Como acionar |
|---|---|---|
| `better-interface` | Revisão completa juntando todas as `better-*` | `/better-interface` |
| `better-ui` | Raios de borda, sombras, ícones, alinhamento óptico | Automática ou pelo nome |
| `better-typography` | Escala tipográfica, espaçamento, quebras, truncamento | Automática ou pelo nome |
| `better-colors` | Cria e confere paletas e tokens semânticos, mede contraste | Automática ou pelo nome |
| `better-accessibility` | Teclado, foco, ARIA, formulários e leitor de tela (WCAG 2.2) | Automática ou pelo nome |
| `better-layout` | Agrupamento, alinhamento, responsividade, espaço para tradução | Automática ou pelo nome |
| `better-writing` | Textos da interface: rótulos, erros, estados vazios | Automática ou pelo nome |
| `interface-review` | Revisa um branch, PR ou mudança local procurando regressões de UI | `/interface-review` |
| `explain-interface` | Explica como um site ou efeito foi feito, a partir de URL ou print | `/explain-interface <url>` |
| `variant` | Gera variantes do componente em que você está trabalhando | `/variant` |
| `state-machine` | Mostra todos os estados de um componente numa página temporária | `/state-machine` |
| `break` | Renderiza um componente em todos os cenários possíveis | `/break` |
| `build-design` | Constrói a UI a partir de imagem ou Figma usando seus tokens | `/build-design` + imagem |

### Limpeza rápida (ibelick)
| Skill | O que é | Como acionar |
|---|---|---|
| `baseline-ui` | Tira a "cara de IA": espaçamento, hierarquia, tipografia | `/baseline-ui` |
| `fixing-accessibility` | Corrige problemas de acessibilidade de forma pontual | `/fixing-accessibility` |
| `fixing-motion-performance` | Corrige animações lentas, layout thrashing e blur pesado | `/fixing-motion-performance` |

### Qualidade, performance e SEO (Addy Osmani)
| Skill | O que é | Como acionar |
|---|---|---|
| `web-quality-audit` | Auditoria geral: performance, acessibilidade, SEO e boas práticas | `/web-quality-audit` |
| `performance` | Velocidade de carregamento e execução | "o site está lento" |
| `core-web-vitals` | Diagnóstico de LCP, INP e CLS | "melhora o LCP" |
| `accessibility` | Conformidade com WCAG 2.2 | "auditoria de acessibilidade" |
| `seo` | Meta tags, dados estruturados, sitemap | "melhora o SEO" |
| `best-practices` | Segurança, compatibilidade, APIs modernas | "aplica boas práticas" |
| `frontend-ui-engineering` | Arquitetura de componentes, design system e acessibilidade | Automática ao construir UI |

### React, Next.js e shadcn
| Skill | O que é | Como acionar |
|---|---|---|
| `web-design-guidelines` | Checa a UI contra mais de 100 regras da Vercel. Boa para rodar antes de publicar | `/web-design-guidelines` |
| `vercel-react-best-practices` | Regras de performance para React e Next | Automática em código React |
| `vercel-composition-patterns` | Componentes compostos, evita excesso de props booleanas | Automática ao refatorar componentes |
| `vercel-react-view-transitions` | Animações com a View Transition API | "transição entre páginas" |
| `shadcn` | Conhece o `components.json` do projeto e o CLI do shadcn | Automática em projetos com shadcn |

### Testes e protótipos
| Skill | O que é | Como acionar |
|---|---|---|
| `webapp-testing` | Abre o app com Playwright, tira prints, clica e lê o console. **Precisa de Python e Playwright** | "testa o fluxo de login no localhost:3000" |
| `web-artifacts-builder` | Protótipos HTML com React, Tailwind e shadcn | "faz um protótipo de…" |
| `taste` | Extrai o "gosto" de design de um site (tokens e decisões). **Precisa do Playwright MCP** | `/taste <url>` |

## Fluxo sugerido

1. **Começar:** `/impeccable init` uma vez por projeto. Depois `/impeccable shape` para planejar e `craft` para construir, ou só peça a tela que `frontend-design` entra automaticamente.
2. **Explorar:** `/prototype` ou `/variant` para comparar versões.
3. **Movimento:** `/find-animation-opportunities` → `/animate` → `/review-animations`.
4. **Polir:** `/baseline-ui` ou `/better-interface`, depois `/impeccable polish`.
5. **Endurecer:** `/break-ui` e `/impeccable harden`.
6. **Antes de publicar:** `/web-design-guidelines` e `/web-quality-audit`.
7. **Antes do PR:** `/interface-review`.

## Origem

Copiadas dos repositórios abaixo em 21/10/2026. Para atualizar, baixe de novo a pasta da skill no repositório de origem.

Observações:
- **Única alteração local:** em `ui-ux-pro-max/SKILL.md`, o caminho `${CLAUDE_PLUGIN_ROOT}/.claude/skills/ui-ux-pro-max` virou `${CLAUDE_SKILL_DIR}`. O original só funciona quando a skill é instalada como plugin. Refaça essa troca se atualizar a skill.
- **Impeccable:** no primeiro uso, o script `impeccable/scripts/impeccable` baixa um executável da página de releases do GitHub do projeto, na versão de `scripts/VERSION`, e confere o hash SHA-256 antes de rodar.
- **Licenças:** as skills da Vercel e a `taste` não trazem arquivo de licença. Ambas se declaram MIT, no frontmatter e no README de origem respectivamente.

| Repositório | Commit | Skills |
|---|---|---|
| [emilkowalski/skills](https://github.com/emilkowalski/skills) | `2d22863` | 13 (sem `write-swift` e `animate-expo`, que são de Swift e React Native) |
| [pbakaus/impeccable](https://github.com/pbakaus/impeccable) | `a0dc1d7` | `impeccable` |
| [Leonxlnx/taste-skill](https://github.com/Leonxlnx/taste-skill) | `6010f91` | `design-taste-frontend` |
| [Leonxlnx/taste-skill-classics](https://github.com/Leonxlnx/taste-skill-classics) | `5e08c93` | `design-taste-frontend-v1`, `high-end-visual-design`, `minimalist-ui`, `industrial-brutalist-ui` |
| [senlindesign/taste-skill](https://github.com/senlindesign/taste-skill) | `6dce223` | `taste` |
| [anthropics/skills](https://github.com/anthropics/skills) | `dbd4588` | `frontend-design`, `webapp-testing`, `web-artifacts-builder`, `theme-factory`, `brand-guidelines` |
| [vercel-labs/agent-skills](https://github.com/vercel-labs/agent-skills) | `063bee9` | `web-design-guidelines`, `vercel-react-best-practices`, `vercel-composition-patterns`, `vercel-react-view-transitions` |
| [shadcn-ui/ui](https://github.com/shadcn-ui/ui) | `ba1503c` | `shadcn` |
| [jakubkrehel/skills](https://github.com/jakubkrehel/skills) | `4c66b1f` | 13 (sem `technical-drawing`, que é desenho técnico de objetos físicos) |
| [ibelick/ui-skills](https://github.com/ibelick/ui-skills) | `55363c3` | `baseline-ui`, `fixing-accessibility`, `fixing-motion-performance` |
| [Dammyjay93/interface-design](https://github.com/Dammyjay93/interface-design) | `2f9be32` | `interface-design` |
| [nextlevelbuilder/ui-ux-pro-max-skill](https://github.com/nextlevelbuilder/ui-ux-pro-max-skill) | `3eccfd8` | `ui-ux-pro-max` |
| [addyosmani/web-quality-skills](https://github.com/addyosmani/web-quality-skills) | `afa8da9` | 6 skills de qualidade web |
| [addyosmani/agent-skills](https://github.com/addyosmani/agent-skills) | `aeac41d` | `frontend-ui-engineering` |
