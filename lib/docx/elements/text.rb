module Docx
  module Elements
    class Text
      include Element

      def self.tag
        't'
      end

      def content
        @node.content
      end

      def content=(args)
        @node.content = args
        # Word/OOXML consumers trim a run's leading/trailing whitespace unless the
        # <w:t> carries xml:space="preserve". Placeholder consolidation and
        # substitution can leave a run holding significant edge whitespace (e.g.
        # " office." after a merged token), so re-assert the attribute to keep the
        # space on render (ENG-3355 "Indiaoffice" collapse).
        if args.is_a?(String) && !args.empty? && args != args.strip
          @node['xml:space'] = 'preserve'
        end
      end

      def initialize(node)
        @node = node
      end
    end
  end
end
