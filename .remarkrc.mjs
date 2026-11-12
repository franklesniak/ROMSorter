export default {
  plugins: [
    [
      'remark-validate-links',
      {
        skipPathPatterns: ['(^|[/\\\\])\\.{3}$']
      }
    ]
  ]
}
